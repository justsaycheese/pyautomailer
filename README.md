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
- **Be able to read `zh_tw` cuz the hardcoding GUI message in python.**
(release has eng version)

Install dependencies with `pip install -r requirements.txt`.
For SMTP mode no additional Windows packages are required.

## Project Structure
```
automailer.py  - entrypoint that launches the GUI
backend.py     - email backends for Outlook or SMTP
gui.py         - Tkinter-based interface and sending logic
utils.py       - helper functions for loading data and generating HTML
```


## Preparing Data
### Recipient List
Create an Excel (`.xlsx`/`.xls`) or CSV file with at least two columns:

- **Email** – recipient address.
- **Salutation** – the greeting used in the message body.

Hidden rows in Excel are ignored.

### Exclusion List (choosable)
Optional Excel/CSV file containing an `Email` column. Any addresses listed
here will be excluded from the send list.

### Message Template
Use an Outlook `.msg` file as the email template. The HTML body can include the
following placeholders which will be replaced when sending:

- `[salutation]` – replaced with the value from the recipient list.
- `[statement]` – replaced with a random closing statement.
- `[image]` – replaced with all selected image (if any images are selected).
- `[image1]`, `[image2]`, ... – inserts a specific image by index

 > When using `[image1]`, `[image2]`, etc., make sure you have enough images loaded.
 >  If an index is out of range, the log will note this issue.

 > If no image selected, the image placeholder will replace with null string.

If the RTF content in the template contains bytes that cannot be decoded,
the program will ignore those bytes to avoid runtime errors.

## Running
Execute the GUI with:

```bash
python automailer.py
```
Or just download the release `.exe` and execute it.

### Interface Features
- Supports Outlook or SMTP
- Allows loading image/attachment folder or selecting multiple files
- Choose between "send" and "save draft" modes

### Settings Persistence
Your settings (accounts, paths, etc.) are saved in `settings.json` and reloaded on next launch.

### Platform Notes
- Outlook mode requires Windows with Outlook installed.
- SMTP mode works cross-platform as long as mail server is reachable.

## Email Sample
```rtf
[salutation]
  body1
[image1]
  body2
[image2]
  body3
[statement]

  ur name
```
