# Automation Workspace

Personal automation repository containing utility scripts, integration tools, and background tasks managed via PowerShell and Windows Task Scheduler.

## Projects

### 📂 syncCalendars
A bidirectional synchronization utility between **Outlook Desktop Calendar** and **Google Calendar**, automated via PowerShell wrappers and registered in Windows Task Scheduler.

- **Language:** Python
- **Automation:** PowerShell (`automate.ps1`, `schedule.ps1`)
- **Task Name:** `mb_syncCalendars`
- **Schedule:** Hourly execution starting at 08:00 AM

## Requirements
- Python 3.x
- Windows 10 / 11
- Microsoft Outlook Desktop client configured with a local profile

## License
This project is licensed under the terms of the [MIT License](LICENSE).