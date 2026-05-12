"""
Async Python wrapper around Microsoft Graph API for clean access to M365.
"""

from gex_msgraph._core import GraphClient
from gex_msgraph._files import FileItem, TreeNode

__all__ = ["GraphClient", "FileItem", "TreeNode"]
__version__ = "0.2.0"
