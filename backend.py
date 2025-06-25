"""Email backends for sending messages."""
import mimetypes
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

try:
    import win32com.client as win32  # type: ignore
except Exception:  # pragma: no cover - non Windows
    win32 = None  # allow import on non-Windows systems

from utils import get_base_dir


class EmailBackend:
    """Base class for email backends."""

    def send(
        self,
        mode: str,
        recipient: str,
        subject: str,
        html_body: str,
        embedded_images: dict[str, Path],
        attachments: list[Path],
    ) -> None:
        raise NotImplementedError


class OutlookBackend(EmailBackend):
    def __init__(self, account_name: str | None = None):
        if win32 is None:
            raise RuntimeError("Outlook backend requires Windows and pywin32")
        self.outlook = win32.Dispatch("Outlook.Application")
        session = self.outlook.GetNamespace("MAPI")
        self.account = None
        if account_name:
            for acct in session.Accounts:
                if acct.DisplayName == account_name:
                    self.account = acct
                    break

    def send(
        self,
        mode: str,
        recipient: str,
        subject: str,
        html_body: str,
        embedded_images: dict[str, Path],
        attachments: list[Path],
    ) -> None:
        mail = self.outlook.CreateItem(0)
        if self.account:
            mail._oleobj_.Invoke(64209, 0, 8, 1, self.account)

        mail.BodyFormat = 2
        mail.Subject = subject
        mail.To = recipient

        for cid, path in embedded_images.items():
            attach = mail.Attachments.Add(
                Source=str(path),
                Type=1,
                Position=0,
            )
            pa = attach.PropertyAccessor
            pa.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
            )
            pa.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True
            )

        mail.HTMLBody = html_body

        for file_path in attachments:
            mail.Attachments.Add(Source=str(file_path))

        if mode == "send":
            mail.Send()
        else:
            mail.Save()


class SmtpBackend(EmailBackend):
    def __init__(self, host: str, port: int, username: str, password: str):
        self.host = host
        self.port = port
        self.username = username
        self.password = password

    def send(
        self,
        mode: str,
        recipient: str,
        subject: str,
        html_body: str,
        embedded_images: dict[str, Path],
        attachments: list[Path],
    ) -> None:
        msg_root = MIMEMultipart("related")
        msg_root["Subject"] = subject
        msg_root["From"] = self.username
        msg_root["To"] = recipient
        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(html_body, "html", "utf-8"))
        msg_root.attach(alt)

        for cid, path in embedded_images.items():
            with open(path, "rb") as f:
                data = f.read()
            mime_type, _ = mimetypes.guess_type(path)
            if mime_type and mime_type.startswith("image/"):
                _, subtype = mime_type.split("/", 1)
            else:
                subtype = path.suffix.lstrip(".") or "png"
            img = MIMEImage(data, _subtype=subtype)
            img.add_header("Content-ID", f"<{cid}>")
            msg_root.attach(img)

        for file_path in attachments:
            with open(file_path, "rb") as f:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition", "attachment", filename=file_path.name
            )
            msg_root.attach(part)

        if mode == "draft":
            draft_dir = get_base_dir() / "drafts"
            draft_dir.mkdir(exist_ok=True)
            with open(draft_dir / f"{recipient}.eml", "w", encoding="utf-8") as f:
                f.write(msg_root.as_string())
            return

        with smtplib.SMTP(self.host, self.port) as server:
            server.starttls()
            server.login(self.username, self.password)
            server.sendmail(self.username, [recipient], msg_root.as_string())
