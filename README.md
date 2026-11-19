# Email Draft Generator

A single-page tool that generates a Python script which creates **Outlook email drafts** from an Excel contact list. Fill in a form, download the script, run it, and review the drafts in Outlook before sending. Nothing is sent automatically.


---

## What it does

- Reads recipients from an Excel file on your Desktop
- Replaces placeholders in the email body (e.g. name, company) with values from each Excel row
- Sets To, CC, BCC, subject and body for every email
- Saves each email to the Outlook **Drafts** folder (never sends)
- Skips rows whose email address is invalid

## Requirements

| Item | Detail |
|------|--------|
| OS | Windows 10 / 11 |
| Mail client | Desktop **Outlook** installed and signed in (classic Outlook; the web version and, in some cases, "New Outlook" are not supported) |
| Python | Python 3 |
| Python packages | `pandas`, `openpyxl`, `pywin32` |
| Excel file | `.xlsx` placed on the Desktop, with a column containing recipient emails |

Install the packages once:

```
python -m pip install pandas openpyxl pywin32
```

## How to use

1. Download `email_generator.html` and open it in a browser.
2. **Excel settings**
   - Enter the Excel file name (e.g. `MyContacts.xlsx`, must be on the Desktop).
   - Enter the name of the email column (e.g. `Email`).
   - Add any other columns you want to use in the email. For each one, give the Excel column name and a placeholder name (e.g. column `Name` with placeholder `姓名`).
3. **Recipients**: optionally fill in CC and BCC (separate multiple addresses with `;`).
4. **Email content**: enter the subject and body. Click a placeholder button to insert it at the cursor, e.g. `Dear {姓名},`.
5. Click **Generate and download Python script** to get `create_drafts.py`.
6. Put `create_drafts.py` on your Desktop, make sure Outlook is open, and run:

```
python C:\Users\<your-username>\Desktop\create_drafts.py
```

7. Open Outlook, go to **Drafts**, check each email, and send manually.

## Example

Excel file:

| Name | Company | Email |
|------|---------|-------|
| Alex Tan | Example Pte Ltd | alex@example.com |

Column mapping in the form:

| Excel column | Placeholder |
|--------------|-------------|
| Name | 姓名 |
| Company | 公司 |

Email body:

```
Dear {姓名},

I would like to introduce an opportunity for {公司}.
```

Resulting draft:

```
Dear Alex Tan,

I would like to introduce an opportunity for Example Pte Ltd.
```

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `pip` or `python` is not recognized | Reinstall Python and tick **Add Python to PATH**, then reopen PowerShell. Use `python -m pip ...` if `pip` alone fails. |
| `KeyError: 'Name'` | The column name in the form does not match the Excel header. Check spelling, case and extra spaces. |
| Excel file not found | Make sure the file is on the Desktop and the file name matches exactly. |
| Outlook error or no drafts appear | Open Outlook first and make sure it is signed in. Switch off "New Outlook" if the script cannot connect. |
| Row skipped | The email cell does not contain a valid address (no `@`). |

## Privacy

The page runs entirely in your browser and never reads or uploads your Excel file. The generated script runs on your own computer and only talks to your local Outlook. Do not commit Excel contact lists or generated scripts containing real email addresses to a public repository.

## Limitations

- Windows only (the script uses `win32com`). It does not work on macOS.
- Creates drafts only; it does not send emails.
- Plain-text email body (no HTML formatting).

---
使用条件：Windows 系统、已安装并登录桌面版 Outlook、已安装 Python 及 `pandas`、`openpyxl`、`pywin32`。
