# Automailer

This project provides a GUI tool for sending batches of emails via Outlook or any SMTP server.

## Requirements
- Python 3.10+
- Packages:
  - `extract_msg`
  - `pandas`
  - `openpyxl`
- `pywin32` (for Outlook mode on Windows)
- `tkinter` (bundled with Python on Windows)

Install dependencies with `pip install -r requirements.txt`.
For SMTP mode no additional Windows packages are required.

## Preparing Data
### Recipient List
Create an Excel (`.xlsx`/`.xls`) or CSV file with at least two columns:

- **Email** – recipient address.
- **Salutation** – the greeting used in the message body.

Hidden rows in Excel are ignored.

### Exclusion List
Optional Excel/CSV file containing an `Email` column. Any addresses listed
here will be excluded from the send list.

### Message Template
Use an Outlook `.msg` file as the email template. The HTML body can include the
following placeholders which will be replaced when sending:

- `[salutation]` – replaced with the value from the recipient list.
- `[statement]` – replaced with a random closing statement.
- `[image]` – replaced with embedded image HTML (if any images are selected).

If the RTF content in the template contains bytes that cannot be decoded,
the program will ignore those bytes to avoid runtime errors.

## Running
Execute the GUI with:

```bash
python automailer_verZ.py
```

Select your recipient list, optional exclusion list, and message template from
the interface. You can also choose images to embed and attachments to include.
Choose **Outlook** or **SMTP** as the backend in the GUI.
For SMTP you must provide the host, port and credentials.
Finally click **Start** to send or draft emails.

### Cross Platform
Outlook mode only works on Windows with Outlook installed.
SMTP mode works on any platform as long as you have network access to your mail server.
