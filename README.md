# Corporate Website Parameter Agent

This agent logs in to a website you are authorized to access and prints visible page parameters such as inputs, selects, textareas, buttons, labels, current values, and select options.
It also creates a work-item field report for pages with fields such as Type, Project Area, Filed Against, Owned By, Planned For, Story Points, Safety, Security, Legal, Description, and Discussion.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
playwright install chromium
```

You can set credentials in your terminal if you want to run without the popup:

```powershell
$env:CORP_USERNAME="your_username"
$env:CORP_PASSWORD="your_password"
```

## Run

Default popup mode:

```powershell
python agent.py
```

The popup asks for:

- Login URL
- Page to inspect, optional
- Username
- Password
- Number of same-domain pages to inspect
- Report file
- Browser, `msedge` or `chromium`
- Whether to show the browser
- Manual login / SSO mode

```powershell
python agent.py --no-popup --login-url "https://example.com/login" --start-url "https://example.com/dashboard" --pages 3
```

Choose a report filename:

```powershell
python agent.py --no-popup --login-url "https://example.com/login" --manual-login --report "ccm_work_item_report.md"
```

Use Microsoft Edge explicitly:

```powershell
python agent.py --no-popup --login-url "https://example.com/login" --manual-login --browser msedge
```

For corporate SSO or MFA login pages:

```powershell
python agent.py --no-popup --login-url "https://example.com/login" --manual-login
```

The browser opens, you complete login yourself, then return to PowerShell and press Enter.

After the run, the agent saves:

- A Markdown report, for example `work_item_report.md`
- A JSON report, for example `work_item_report.json`

For debugging:

```powershell
python agent.py --no-popup --login-url "https://example.com/login" --show-browser
```

If the site uses unusual login fields, edit the selectors inside `login()` in `agent.py`.
