"""
Async Python wrapper around Microsoft Graph API for clean access to M365.
"""

from gex_msgraph._core import GraphClient
from gex_msgraph._exceptions import (
    AuthError,
    GraphAuthenticationError,
    GraphError,
    GraphSyncInLoopError,
    NotFoundError,
    PermissionDeniedError,
    RateLimitExhaustedError,
)
from gex_msgraph._files import DriveLocation, FileItem, TreeNode

__all__ = [
    "GraphClient",
    "FileItem",
    "DriveLocation",
    "TreeNode",
    "GraphError",
    "AuthError",
    "PermissionDeniedError",
    "NotFoundError",
    "RateLimitExhaustedError",
    "GraphAuthenticationError",
    "GraphSyncInLoopError",
]
__version__ = "0.5.0"
