GoogleAppSuicideSquad
A powerful Google Apps Script automation tool for tracking changes in Google Sheets , logging activity history, and exporting data to reports — all integrated with Google Drive . 

This project is designed to help manage and monitor data in a military-style unit (battalion or division), ensuring full transparency of user actions, maintaining logs, and generating structured reports when needed.

🧰 Features
✅ Change Logging
Tracks additions, deletions, edits, and formula updates on key sheets.
Logs are stored in a dedicated sheet ("Лог змін" / "Change Log").
Different color highlighting for change types:
Green: New entries
Red: Deletions
Blue: Edits
Yellow: Formula changes
📤 Data Exporting
Export logs and data to:
.xlsx (Excel)
.csv
Optional export templates for Microsoft Word via custom dialogs or third-party integrations.
Auto-cleanup of old backups based on retention settings.
📁 Google Drive Integration
Stores temporary files, archives, and logs in a designated folder on Google Drive (via TMP_FOLDER_ID).
Ability to list and manage all log files directly from the Drive.
🗂 Archive Management
Manual and automated archiving of logs by date.
Supports daily triggers for automatic cleanup and backup.
🔍 Validation & Analytics
Spell check and format validation to ensure consistent data entry.
User-based analytics: see who edited what, and how often.
🧭 Custom Menu UI
Adds a custom menu in Google Sheets for quick access to:
Export options
Spell check
Manual validation
Activity reports
Trigger management

📁 File Structure
Here’s an overview of the files in your project:

1. general.gs
Contains core configuration and setup logic.
Defines constants like sheet names, important ranges, and color codes.
backupLogSheet.gs
profile
Qwen3-235B-A22B
11:25 am
Отлично! Теперь, когда я вижу структуру ваших файлов и кода, я могу улучшить описание проекта на GitHub. Я буду писать описание на английском , а комментарии для вас (на русском) будут помечены как "💡 Примечание" .

🎯 GoogleAppSuicideSquad
A powerful Google Apps Script automation tool for tracking changes in Google Sheets , logging activity history, and exporting data to reports — all integrated with Google Drive . 

This project is designed to help manage and monitor data in a military-style unit (battalion or division), ensuring full transparency of user actions, maintaining logs, and generating structured reports when needed.

🧰 Features
✅ Change Logging
Tracks additions, deletions, edits, and formula updates on key sheets.
Logs are stored in a dedicated sheet ("Лог змін" / "Change Log").
Different color highlighting for change types:
Green: New entries
Red: Deletions
Blue: Edits
Yellow: Formula changes
📤 Data Exporting
Export logs and data to:
.xlsx (Excel)
.csv
Optional export templates for Microsoft Word via custom dialogs or third-party integrations.
Auto-cleanup of old backups based on retention settings.
📁 Google Drive Integration
Stores temporary files, archives, and logs in a designated folder on Google Drive (via TMP_FOLDER_ID).
Ability to list and manage all log files directly from the Drive.
🗂 Archive Management
Manual and automated archiving of logs by date.
Supports daily triggers for automatic cleanup and backup.
🔍 Validation & Analytics
Spell check and format validation to ensure consistent data entry.
User-based analytics: see who edited what, and how often.
🧭 Custom Menu UI
Adds a custom menu in Google Sheets for quick access to:
Export options
Spell check
Manual validation
Activity reports
Trigger management
📁 File Structure
Here’s an overview of the files in your project:

1. general.gs
Contains core configuration and setup logic.
Defines constants like sheet names, important ranges, and color codes.
profile
Qwen3-235B-A22B
11:25 am
Отлично! Теперь, когда я вижу структуру ваших файлов и кода, я могу улучшить описание проекта на GitHub. Я буду писать описание на английском , а комментарии для вас (на русском) будут помечены как "💡 Примечание" .

🎯 GoogleAppSuicideSquad
A powerful Google Apps Script automation tool for tracking changes in Google Sheets , logging activity history, and exporting data to reports — all integrated with Google Drive . 

This project is designed to help manage and monitor data in a military-style unit (battalion or division), ensuring full transparency of user actions, maintaining logs, and generating structured reports when needed.

🧰 Features
✅ Change Logging
Tracks additions, deletions, edits, and formula updates on key sheets.
Logs are stored in a dedicated sheet ("Лог змін" / "Change Log").
Different color highlighting for change types:
Green: New entries
Red: Deletions
Blue: Edits
Yellow: Formula changes
📤 Data Exporting
Export logs and data to:
.xlsx (Excel)
.csv
Optional export templates for Microsoft Word via custom dialogs or third-party integrations.
Auto-cleanup of old backups based on retention settings.
📁 Google Drive Integration
Stores temporary files, archives, and logs in a designated folder on Google Drive (via TMP_FOLDER_ID).
Ability to list and manage all log files directly from the Drive.
🗂 Archive Management
Manual and automated archiving of logs by date.
Supports daily triggers for automatic cleanup and backup.
🔍 Validation & Analytics
Spell check and format validation to ensure consistent data entry.
User-based analytics: see who edited what, and how often.
🧭 Custom Menu UI
Adds a custom menu in Google Sheets for quick access to:
Export options
Spell check
Manual validation
Activity reports
Trigger management
📁 File Structure
Here’s an overview of the files in your project:

1. general.gs
Contains core configuration and setup logic.
Defines constants like sheet names, important ranges, and color codes.
2. backupLogSheet.gs
Handles log archiving and cleanup.
Manages temporary files and folders on Google Drive.
3. WordExportForm.html
Provides a custom dialog for exporting logs to Microsoft Word.
Uses HTML/CSS/JavaScript for interactive forms.
4. logs_restore.gs
Restores archived logs back into the main spreadsheet.
Useful for recovering historical data.
5. globals.gs
Contains global variables and utility functions used across scripts.
6. logs_restor.html
Another custom dialog for restoring logs, likely with form inputs.
7. exportToWordWithDialog.gs
Integrates with WordExportForm.html to handle exports to Word.
8. export_To_Word.gs
Core logic for exporting data to Microsoft Word.
9. validator.gs
Performs data validation and spell checking.
Ensures consistency in data entry.
10. user_report.gs
Generates user activity reports based on logged changes.
