# GoogleAppSuicideSquad

A powerful Google Apps Script automation suite for Google Sheets — enabling advanced change tracking, audit logging, analytics, and seamless exports to Excel, CSV, and Word, all fully integrated with Google Drive.  
Originally designed for military-style units (battalions/divisions), this toolkit is perfect for any team or organization needing transparency, robust data control, and effortless reporting.

---

# GoogleAppSuicideSquad

Потужний набір інструментів автоматизації Google Apps Script для Google Sheets — забезпечує просунуте відстеження змін, аудит, аналітику та безшовний експорт в Excel, CSV і Word, повністю інтегрований з Google Drive.  
Спочатку розроблений для підрозділів типу логістики, інформаційних технології, бухгалтерії, цей інструментарій ідеально підходить для будь-якої команди або організації, що потребує прозорості, надійного контролю даних та легкого створення звітів.

---

## 🚀 Features

### ✅ Change Logging & Audit Trail
- **Tracks all edits**: additions, deletions, value edits, and formula changes across key sheets.
- **Detailed logs**: Every change is recorded in a dedicated "Change Log" sheet, including timestamp, user, sheet, cell, action type, before/after values, formulas, and "important" flag.
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

## 🚀 Можливості

### ✅ Система аудиту та відстеження змін
- **Відстежує всі зміни**: додавання, видалення, редагування значень та формул у всіх ключових листах.
- **Детальні логи**: Кожна зміна записується в спеціальний лист "Лог змін" з указанням часу, користувача, листа, комірки, типу дії, старих/нових значень та прапорця "важливості".
- **Візуальний зворотний зв'язок**: Кольорова індикація змін (зелений - додано, червоний - видалено, жовтий - змінено).

### 📁 Інтеграція з Google Drive та архівування
- **Централізоване зберігання**: Всі логи, тимчасові файли та експорти зберігаються в призначеній папці Google Drive.
- **Гнучий експорт**: Одним кліком експортуйте логи/дані в Excel (.xlsx), CSV або Word (.docx) — з прямими посиланнями для завантаження.
- **Автоматичне архівування**: Щоденне або ручне архівування логів з очищенням старих резервних копій на основі налаштувань зберігання.
- **Управління файлами**: Швидке перерахування, відновлення або видалення резервних/архівних файлів.

### 📤 Просунутий експорт та звітність
- **Експорт в Word**: Миттєво експортуйте будь-який лист або діапазон в Word з налаштовуваними діалогами для складних звітів (заголовки, множинні таблиці, описи).
- **Користувацькі діалоги**: HTML-спливаючі вікна та бічні панелі для вибору опцій експорту, форматування та попереднього перегляду.
- **Генератор користувацьких звітів**: Заповнюйте форми для створення структурованих Word-звітів з динамічними таблицями та метаданими.

### 🔍 Валідація, пошук та аналітика
- **Валідація даних**: Забезпечує узгоджений, якісний ввід даних (перевірка форматів, орфографії тощо).
- **Пошук по історії**: Потужний пошук та витяг логів змін та історичних даних.
- **Аналітика користувачів**: Розумійте, хто що змінив, коли та де — зі звітами про діяльність за користувачами, листами та датами.

### 🧭 Користувацький інтерфейс та автоматизація
- **Розумне меню**: Додає користувацьке меню в Google Sheets для миттєвого доступу до всіх основних функцій (аудит, експорт, форматування, тригери тощо).
- **Швидке форматування**: Легко налаштовуйте висоту рядків через бічну панель.
- **Автоматичні тригери**: Плануйте щоденні резервні копії та архівування без ручного втручання.

---

## 📦 File Structure

```
/
├── menu.js                     # Основне меню, синхронізація між підрозділами
├── log.js                      # Система логування змін та аудиту
├── user_report.js              # Аналітика дій користувачів та звіти
├── get_keys_flexble.js         # Гнуча генерація ключів та QR-кодів
├── globals.js                  # Глобальні змінні та утиліти
├── global_table_search.js      # Глобальний нечіткий пошук по всіх листах
├── exportToWordWithDialog.js   # Просунутий експорт в Word (діалоги, множинні таблиці)
├── export_To_World.js          # Простий експорт діапазону листа в Word
├── SidebarG                    # Скрипт для логіки UI висоти рядків
├── Sidebar.html                # HTML бічна панель для форматування
├── WordExportForm.html         # HTML форма для складних Word-звітів
├── GlobalFuzzySearch.html      # HTML діалог для глобального пошуку
├── README.md                   # Ви тут!
├── CONTRIBUTING.md             # Керівництво по участі в проекті
└── LICENSE                     # Ліцензія MIT
```

---

## 🛡️ Use Cases

- **Military units**: Track equipment, personnel, and supply changes with full auditability.
- **Business teams**: Ensure transparency and accountability in collaborative spreadsheets.
- **Project management**: Maintain a tamper-proof history and generate professional reports on demand.
- **Any organization**: Where change tracking, compliance, and reliable backup matter.

---

## 🛡️ Застосування

- **Підрозділи 1РБпАК, 2ББпАК, 3РБпАК**: Відстеження обладнання, персоналу та змін в постачанні з повною аудируваністю.
- **Бізнес-команди**: Забезпечення прозорості та підзвітності в спільних таблицях.
- **Управління проектами**: Підтримка захищеної від змін історії та генерація професійних звітів за запитом.
- **Будь-яка організація**: Де важливі відстеження змін, відповідність вимогам та надійне резервне копіювання.

---

## ⚡ Quick Start

1. **Copy all scripts/files** into your Google Apps Script project attached to your Google Sheet.
2. **Set your Google Drive folder ID** (`TMP_FOLDER_ID`) for backups/archives.
3. Reload your Google Sheet — the custom menu will appear automatically.
4. Start tracking, analyzing, exporting, and feeling secure!

---

## ⚡ Швидкий старт

1. **Скопіюйте всі скрипти/файли** в ваш проект Google Apps Script, прикріплений до Google Таблиці.
2. **Встановіть ID папки Google Drive** (`TMP_FOLDER_ID`) для резервних копій/архівів.
3. Перезавантажте Google Таблицю — користувацьке меню з'явиться автоматично.
4. Почніть відстежувати, аналізувати, експортувати та почуватися в безпеці!

---

## 📋 Credits & License

Developed by [Dmitze](https://github.com/Dmitze).  
MIT License.  
Contributions and feedback are welcome!

---

## 📋 Автори та ліцензія

Розроблено [Дмитрієм Шивачовим](https://t.me/Dmitry_Shiva).  
MIT License.  
Внесок у проект та зворотній зв'язок вітаються!

---
