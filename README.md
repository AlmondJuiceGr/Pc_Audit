# pc_audit

Enterprise desktop audit dashboard for collecting workstation inventory, running Windows device checks, and viewing manual test results from network clients.

## Main Files

- `enterprise_tabs_app.pyw` - Tkinter desktop app.
- `enterprise_tabs_app.py` - console launcher for the `.pyw` app.
- `network_job_server.py` - local network job/client registration server.
- `register_client.py` - client inventory and diagnostics reporter.
- `*.bat` / `*.sh` - helper launch scripts for Windows and Linux.

Runtime databases, exports, cache files, and logs are intentionally ignored by Git.
