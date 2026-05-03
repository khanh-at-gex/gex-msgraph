"""
Async Python wrapper around Microsoft Graph API for clean access to M365.
"""

from gex_msgraph._core import GraphClient
from gex_msgraph._files import FileItem

__all__ = ["GraphClient", "FileItem"]
__version__ = "0.1.0"
