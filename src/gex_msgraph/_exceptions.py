from __future__ import annotations

from typing import NoReturn

import httpx


class GraphError(httpx.HTTPStatusError):
    """Base error for Microsoft Graph API failures.

    Subclasses ``httpx.HTTPStatusError`` so existing ``except
    httpx.HTTPStatusError`` call sites keep working unchanged. The original
    request/response stay available via ``.request`` / ``.response``.
    """


class AuthError(GraphError):
    """401 Unauthorized — token invalid, expired, or insufficient."""


class PermissionDeniedError(GraphError):
    """403 Forbidden — authenticated but not authorized for this resource."""


class NotFoundError(GraphError):
    """404 Not Found."""


class RateLimitExhaustedError(GraphError):
    """429/5xx persisted through every retry attempt."""


class GraphAuthenticationError(Exception):
    """Token acquisition failed before any Graph HTTP call was made.

    Separate root from GraphError because MSAL failures carry no
    httpx request/response to attach.
    """


class GraphSyncInLoopError(RuntimeError):
    """A *_sync method was called from within a running event loop."""


_STATUS_MAP: dict[int, type[GraphError]] = {
    401: AuthError,
    403: PermissionDeniedError,
    404: NotFoundError,
}


def raise_graph_error(
    response: httpx.Response, *, retries_exhausted: bool = False
) -> NoReturn:
    """Raise the GraphError subclass matching the response's status code."""
    status = response.status_code
    if retries_exhausted:
        cls: type[GraphError] = RateLimitExhaustedError
    else:
        cls = _STATUS_MAP.get(status, GraphError)
    message = (
        f"{status} error for {response.request.method} '{response.request.url}'"
    )
    raise cls(message, request=response.request, response=response)
