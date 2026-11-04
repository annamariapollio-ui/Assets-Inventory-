# VI Rollout · Market Inventory survey app

One codebase, one link per market. Contributors answer the 33 touchpoints (workstreams A–F) in the browser;
on submission the answers are written into the existing Excel template (Inventory tab, columns C–O).

## Run locally
    pip install openpyxl          # only dependency, Python 3.9+
    python3 server.py             # prints the admin link
    python3 server.py add "Switzerland" CHF    # prints the market link (also possible from the admin page)

Default port 8080 (`PORT=9000 python3 server.py` to change). Set `PUBLIC_URL` so printed links carry the real domain.

## How it works
- **Market link** `/m/<market>?k=<key>`: market name and currency are pre-set; the key controls access.
- **Workstream cards**: each workstream has a default owner and a "Copy link" button for handing it to a colleague.
- **Touchpoint form**: starts with "Does it exist?"; follow-up questions appear only on "Yes". Assets and identities are captured one row per asset, quantity as exact count / band / not applicable, cost in the market currency.
- **Live collaboration**: autosave, presence badges ("Sabrina active"), and a prompt when two people edit the same touchpoint.
- **Submit gate**: opens only when all 33 touchpoints are answered ("Unknown" counts) and all 6 workstream owners are set (`REQUIRE_OWNERS` in server.py).
- **Output**: `submissions/<market>/<market>_<timestamp>.xlsx`, downloadable from `/admin?a=<admin key>`.
  Optional e-mail delivery: set `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `MAIL_FROM`, `MAIL_TO`.
- **Template**: `template/NEW_PMI_Market_Quantitative_Workbook.xlsx`. Replacing the file (same layout) updates the export; touchpoints are read from it at start-up.

## Before production use (for alignment with PMI IT / security)
- Host inside an approved environment (Azure App Service, container, or internal server) behind HTTPS; the market key is a lightweight access control and can be replaced by SSO.
- Data is stored as JSON files in `data/`; a managed database or SharePoint/Blob storage is the natural next step.
- The Recap tab of the template still contains broken `#REF!` links from the previous structure and is left untouched by the export.
