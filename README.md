# GNC MIS automation

Daily refresh of the GNC dashboards (Excel `.xlsb`, not stored here - they hold customer data).
Everything runs from `D:\MIS\GNC\Automate` on Windows with Excel, Python 3 + Selenium and Chrome.

| Launcher | What it does |
|---|---|
| `GNC_Inbound.bat` | Downloads Feedback (244), CDR Report IN (240) and DialDesk In Call Details, pastes them into the **Inbound MIS Dashboard** (`Feedback Raw`, `CDR Raw`, `Inbound DD Tagging`), drags formulas, refreshes pivots, shows the days up to yesterday on `Sales Performance`. Scheduled 07:40. |
| `GNCchatmis.bat` | KwikEngage ticket export (Analytics > Export Reports > mail link) -> `Chat` sheet of the **Chat Dashboard**; shows the days up to yesterday on `Snap`. Rolling 10-day window. Scheduled 07:55. |
| `GNC_Chat_Inbound.bat` | Copies the Chat / Inbound sales from the Overall Sale dashboard into the Chat (`Sale Raw`) and Inbound (`Sales Raw`) dashboards and appends the day's APR. |
| `scripts/GNC_Auto.bat` | Builds the Overall Sale Performance dashboard (Rohan's mail, CDR, allocation, APR, pivots). |

Each `.bat` calls the `.ps1` of the same name (Excel automation through COM) and, where needed, a `fetch_*.py` downloader.
Every paste is read back and compared with the source; nothing is saved if anything differs. Re-running is safe (no duplicates).
Keep the workbook closed while a job runs.

## Setup
1. Copy `inbound.env.example` -> `inbound.env` and `chat.env.example` -> `chat.env`; fill in the logins. These files are git-ignored.
2. Gmail (read-only IMAP) login for the mail steps: the `email.env.txt` path used in `fetch_chat.py` / `scripts/fetch_rohan.py`.
3. `pip install selenium webdriver-manager requests`
