# GoogleAppSuicideSquad

A powerful Google Apps Script automation suite for Google Sheets — enabling advanced change tracking, audit logging, analytics, and seamless exports to Excel, CSV, and Word, all fully integrated with Google Drive.  
Originally designed for military-style units (battalions/divisions), this toolkit is perfect for any team or organization needing transparency, robust data control, and effortless reporting.

---

## 🚀 Features

### ✅ Change Logging & Audit Trail
- **Tracks all edits**: additions, deletions, value edits, and formula changes across key sheets.
- **Detailed logs**: Every change is recorded in a dedicated “Change Log” sheet, including timestamp, user, sheet, cell, action type, before/after values, formulas, and “important” flag.
- **Visual feedback**: Color-coded highlighting of changes (green for additions, red for deletions, yellow for edits).

### 📁 Google Drive Integration & Archiving
- **Centralized storage**: All logs, temporary files, and exports are kept in a designated Google Drive folder.
- **Flexible exports**: One-click export of logs/data to Excel (.xlsx), CSV, or Word (.docx) — with direct download links.
- **Automatic archiving**: Daily or manual log archiving with cleanup of old backups based on retention settings.
- **File management**: Quickly list, restore, or delete backup/archive files.

### 📤 Advanced Exporting & Reporting
- **Export to Word**: Instantly export any sheet or range to Word, with custom dialogs for complex reports (headers, multiple tables, descriptions).
- **User-friendly dialogs**: HTML-based popups and sidebars for choosing export options, formatting, and previews.
- **Custom report generator**: Fill out forms to create structured Word reports with dynamic tables and metadata.

### 🔍 Validation, Search, and Analytics
- **Data validation**: Ensures consistent, high-quality data entry (format checks, spell checks, etc.).
- **History search**: Powerful search and retrieval of change logs and historical data.
- **User analytics**: Understand who changed what, when, and where — with activity reports by user, sheet, and date.

### 🧭 Custom UI & Automation
- **Smart menu**: Adds a custom menu in Google Sheets for instant access to all major features (audit, export, formatting, triggers, etc).
- **Quick formatting**: Easily adjust row heights via sidebar interface.
- **Automated triggers**: Schedule daily backups and archiving without manual intervention.

---

## 📦 File Structure

```
/
├── general.gs                  # Core logic, configuration, constants, menu, change monitoring
├── backupLogSheet.js           # Log backup/archiving, cleanup, Drive file management
├── export_To_World.js          # Simple export of sheet range to Word
├── exportToWordWithDialog.js   # Advanced Word export (custom dialog, multi-table, headers)
├── validator.gs                # Data validation and checks
├── user_report.gs              # User activity analytics and reports
├── logs_restore.gs             # Restoration of archived logs
├── globals.gs                  # Global variables and utility functions
├── SidebarG                    # Script for row height UI logic
├── Sidebar.html                # HTML sidebar for formatting
├── WordExportForm.html         # HTML form for complex Word reports
├── logs_restor.html            # HTML dialog for log restoration
├── README.md                   # You are here!
```

---

## 🛡️ Use Cases

- **Military units**: Track equipment, personnel, and supply changes with full auditability.
- **Business teams**: Ensure transparency and accountability in collaborative spreadsheets.
- **Project management**: Maintain a tamper-proof history and generate professional reports on demand.
- **Any organization**: Where change tracking, compliance, and reliable backup matter.

---

## ⚡ Quick Start

1. **Copy all scripts/files** into your Google Apps Script project attached to your Google Sheet.
2. **Set your Google Drive folder ID** (`TMP_FOLDER_ID`) for backups/archives.
3. Reload your Google Sheet — the custom menu will appear automatically.
4. Start tracking, analyzing, exporting, and feeling secure!

---

## 📋 Credits & License

Developed by [Dmitze](https://github.com/Dmitze).  
MIT License.  
Contributions and feedback are welcome!

---
