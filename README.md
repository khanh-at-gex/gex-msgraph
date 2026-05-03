# gex-msgraph

Async Python wrapper around Microsoft Graph API for clean access to M365 (SharePoint, OneDrive, Outlook, Teams) from any context — FastAPI, Prefect, scripts, notebooks.

**Status:** v0.1.0 | Python >=3.11 | Internal/Private

## Quickstart

```python
from gex_msgraph import GraphClient
import asyncio

async def main():
    # Credentials loaded from MS_DAS_U1_* env vars automatically
    async with GraphClient("das_u1") as client:
        # Read Excel directly to a pandas DataFrame
        df = await client.read_excel(item_path="Reports/Q1.xlsx")
        print(df.head())

if __name__ == "__main__":
    asyncio.run(main())
```

## Documentation

- **[USAGE.md](USAGE.md)** — the main end-user guide: installation, setup, recipes, and deployment.
- **[AGENTS.md](AGENTS.md)** — internal documentation for AI coding assistants.
- **[CHANGELOG.md](CHANGELOG.md)** — version history.
