# gex-msgraph — Capability Map

```
gex-msgraph
├── Setup
│   ├── GraphClient(account, *, client_id, client_secret, tenant_id,
│   │               username, password, default_drive_id,
│   │               max_concurrent, request_timeout)
│   ├── Credentials  MS_<ACCOUNT>_CLIENT_ID / SECRET / TENANT_ID
│   │                MS_<ACCOUNT>_USERNAME / PASSWORD  (optional pair)
│   ├── Identifiers  item_path | share_url | item_id  (exactly one)
│   ├── Auth         delegated ROPC (username+password) or
│   │                app-only client credentials (omit both) · auto-retry 429 / 5xx
│   └── Errors       GraphError(httpx.HTTPStatusError) → AuthError · PermissionDeniedError
│                    · NotFoundError · RateLimitExhaustedError
│                    GraphAuthenticationError · GraphSyncInLoopError
│
├── File Operations
│   ├── Reading
│   │   ├── read_excel(*, item_path|share_url|item_id, sheet=0, **kwargs)
│   │   ├── read_csv(*, item_path|share_url|item_id, **kwargs)
│   │   ├── read_parquet(*, item_path|share_url|item_id, **kwargs)
│   │   ├── read_excel_many(paths, *, sheet, sheet_match, on_missing_sheet,
│   │   │                   on_error, add_source_column, return_status, **kwargs)
│   │   ├── read_csv_many(paths, *, on_error, add_source_column,
│   │   │                 return_status, **kwargs)
│   │   ├── list_excel_sheets(*, item_path|share_url|item_id)   xlsx workbooks only
│   │   └── SpreadsheetML 2003  auto-detected + decoded by read_excel(_many)
│   │                           (SAP ".xls" = UTF-16 XML) · values only
│   │
│   ├── Discovery
│   │   ├── list_files(folder_path)           immediate children, files + folders
│   │   ├── walk(folder_path, *, pattern, recursive)   files only, fnmatch glob
│   │   ├── search_files(query, *, limit)     search by name / content across drive
│   │   ├── get_metadata(*, item_path|share_url|item_id)   → FileItem
│   │   ├── exists(*, item_path|share_url|item_id)         → bool
│   │   └── get_folder_tree(folder_path)      → TreeNode  (.print())
│   │
│   ├── Management
│   │   ├── download(*, item_path|share_url|item_id, to_path=None) → bytes | Path
│   │   ├── upload(local_path, remote_path)             → dict  (>4 MiB chunked)
│   │   ├── upload_many([(local, remote), ...], *, on_error,
│   │   │               max_concurrent, return_status)
│   │   ├── delete_file(*, item_path|share_url|item_id)
│   │   ├── copy_file(*, item_path|share_url|item_id, dest_folder_path,
│   │   │             new_name, wait, wait_timeout)     → FileItem | None
│   │   ├── move_file(*, item_path|share_url|item_id,
│   │   │             dest_folder_path, new_name)       → FileItem
│   │   ├── create_folder(folder_path)                  → FileItem
│   │   └── get_share_link(*, item_path|share_url|item_id,
│   │                       link_type, scope)           → str (webUrl)
│   │
│   └── Sync Wrappers  (background loop — scripts / no event loop)
│       ├── read_excel_sync(**kwargs)
│       ├── read_csv_sync(**kwargs)
│       ├── download_sync(**kwargs)
│       ├── read_excel_many_sync(paths, **kwargs)
│       ├── read_csv_many_sync(paths, **kwargs)
│       └── close_sync()
│
├── Communication
│   ├── Email
│   │   ├── list_mail(limit, *, folder)       → list[dict]
│   │   └── send_mail(to, subject, body, *, body_type, cc, attachments)
│   ├── Teams — Channel
│   │   ├── send_teams_message(team_id, channel_id, text)
│   │   └── get_teams_messages(team_id, channel_id, limit)  → list[dict]
│   └── Teams — Chat
│       ├── list_chats(limit)             → list[dict]  (1-1 · group · meeting)
│       ├── get_chat_messages(chat_id, limit)  → list[dict]
│       └── send_chat_message(chat_id, text)
│
└── Public Types
    ├── FileItem    name · path · id · size · modified · is_folder · webUrl
    └── TreeNode    item: FileItem | None · children: list[TreeNode] · print()
```
