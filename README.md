# Automailer

This project provides a GUI tool for sending batches of emails via Outlook.

## Requirements
- Python 3.10+
- Packages:
  - `extract_msg`
  - `pandas`
  - `openpyxl`
  - `pywin32` (provides `win32com.client`)
  - `tkinter` (bundled with Python on Windows)

Install dependencies with `pip install extract_msg pandas openpyxl pywin32`.

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

## Running
Execute the GUI with:

```bash
python automailer_verZ.py
```

Select your recipient list, optional exclusion list, and message template from
the interface. You can also choose images to embed and attachments to include.
Finally click **Start** to send or draft emails using Outlook.

## Building an Executable
To package the program into a standalone Windows executable and prepare a release archive, run:

```bash
python build_release.py
```

The script relies on `pyinstaller`. Install it first with `pip install pyinstaller`.
When complete, `release.zip` will contain `automailer_verZ.exe`, the default
`embed` and `attachment` folders, and the example `sample.msg` file.
