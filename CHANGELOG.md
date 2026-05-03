# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [0.1.0] - 2026-05-03

### Added
- Async Python wrapper around Microsoft Graph API for M365 access.
- `GraphClient` public class with ROPC auth flow and connection management.
- `FileItem` dataclass for SharePoint/OneDrive file representation.
- Read files single/bulk: `read_excel`, `read_csv`, `download`, `read_excel_many`.
- Discovery endpoints: `walk`, `list_files`.
- Write endpoint: `upload`.
- Communication endpoints: `send_mail`, `send_teams_message`.
- Sync wrappers for ease of use in scripts/notebooks: `read_excel_sync`, `read_csv_sync`, `download_sync`.
- Automatic retry logic and concurrency capping per account.
