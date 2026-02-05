# WhatsApp Bulk Messenger

Simple script to send WhatsApp messages from an Excel file using Selenium.

## Quick Start

1. Install dependencies:

```powershell
pip install -r requirements.txt
```

If you have Python 3.13 and experience install issues, run:

```powershell
pip install selenium pandas openpyxl webdriver-manager
```

2. Edit `contacts.xlsx` — two columns: `phone` and `message` (phone must include country code, no `+` or spaces).

3. Run the script:

```powershell
python send_whatsapp.py
```

Or double-click `RUN_ME.bat` on Windows.

## Notes
- The script auto-downloads the appropriate ChromeDriver via `webdriver-manager`.
- On first run you will need to scan the WhatsApp Web QR code. A `whatsapp_session` folder is used to persist login.
- If Chrome fails to start, ensure Chrome is installed and matches the driver; update Chrome if necessary.

## Troubleshooting
- If you see issues installing packages, try upgrading pip: `pip install --upgrade pip`.
- If Chrome shows a DevToolsActivePort error, the script will retry using a temporary profile; ensure no other Chrome instance is blocking startup.

## Files
- `send_whatsapp.py` — main script
- `contacts.xlsx` — contacts and messages
- `requirements.txt` — Python dependencies
- `RUN_ME.bat`, `run_me.sh` — launchers

## License & Privacy
This runs locally only. No data is sent to external servers.

---

If you'd like, I can also:
- add a `--dry-run` mode, or
- update `RUN_ME.bat` to call `python -u send_whatsapp.py` for clearer console output.
