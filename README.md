# Job Application Tracker – Google Sheets

A Google Apps Script tool that turns a Google Sheets spreadsheet into a lightweight application tracker, with a custom menu and a dynamic HTML form for logging new applications directly from the sheet.

## Demo

_Add a screenshot or short GIF here showing the "Applications" menu and the submission form._

## Features

- **Custom menu** — Adds an "Applications" menu to Google Sheets for quick access to the tool.
- **Dynamic form** — An HTML interface to log a new application, with a dropdown to pick an existing company or add a new one on the fly.
- **Automation** — New applications are appended to the sheet automatically, dropdown lists are kept up to date, and rows are conditionally formatted (color-coded) by company for easier scanning.

## Scripts

| File | Purpose |
|---|---|
| `Code.gs` | Core logic: adds the custom menu, appends new applications to the sheet, manages the company dropdown list, and applies conditional formatting. |
| `FormulaireCandidature.html` | The form UI shown to the user — company, job title, and link fields, with input validation and dynamic company selection. |

## Getting started

1. **Get the template spreadsheet**
   Duplicate the [example spreadsheet](https://docs.google.com/spreadsheets/d/11Sz7P_bk9WTBEnWX4phSnCvWJVZwe-19hquaYakiXXw/edit?usp=sharing) into your own Google Drive, or set up your own sheet with an equivalent structure.

2. **Add the scripts**
   In your spreadsheet, go to `Extensions > Apps Script`, then create a script file for each file in this repo (`Code.gs` and `FormulaireCandidature.html`) and paste its content in.

3. **Use it**
   Reload the spreadsheet — an "Applications" menu appears automatically. Click it, choose "Add an application", fill in the company, job title, and link, then submit. The entry is added to the sheet and the company list/formatting update automatically.

## Tech stack

- Google Apps Script (JavaScript)
- HTML/CSS (form UI)
- Google Sheets

## Author

Developed by [Clément Vongsanga](https://github.com/Synergy-XVortex). Questions, suggestions, and pull requests are welcome.
