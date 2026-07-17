import importlib.metadata

import gex_msgraph


def test_installed_version_matches_package_version():
    assert importlib.metadata.version("gex-msgraph") == gex_msgraph.__version__
