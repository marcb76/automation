# Automation Workspace

Personal automation repository containing utility scripts, integration tools, and background tasks managed via PowerShell and Windows Task Scheduler.

## Projects

### 📂 syncCalendars
A bidirectional synchronization utility between **Outlook Desktop Calendar** and **Google Calendar**, automated via PowerShell wrappers and registered in Windows Task Scheduler.

- **Language:** Python
- **Automation:** PowerShell (`automate.ps1`, `schedule.ps1`)
- **Task Name:** `mb_syncCalendars`
- **Schedule:** Hourly execution starting at 08:00 AM

### 📂 searchAndExportEmails
An automated utility designed to recursively search, filter, and export emails from local Outlook desktop PST stores based on target addresses, date ranges, folder regex, and subject criteria.

- **Language:** Python
- **Core Features:** 
  - Isolated temporary workspace processing.
  - Automatic conversion of emails to a clean PDF format (`_email.pdf`) positioned at the top of each folder.
  - Intelligent extraction of true standalone attachments while filtering out Outlook auto-generated signature/inline graphics (`image*.png`, etc.).
  - Dedicated per-email chronological folder structure (`YYYYMMDD_HHMMSS_Subject`) to ensure flawless sorting in Windows Explorer.
  - Robust file-locking error handling with interactive recovery prompts.

> **⚠️ Configuration Note:** The scripts reference absolute local paths (e.g., `C:\Users\marcb\...`). If you clone or use this repository, you must update the target paths to match your own local directory structure before execution.

## Requirements
- Python 3.x
- Windows 10 / 11
- Microsoft Outlook Desktop client configured with a local profile
- Microsoft Word (required for HTML-to-PDF email conversions via COM automation)

## License
This project is licensed under the terms of the [MIT License](LICENSE).