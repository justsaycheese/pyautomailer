import os
import random
import re
import time
import pythoncom
from tkinter import Tk, messagebox

import extract_msg
import pandas as pd

from backend import OutlookBackend, SmtpBackend
from utils import (
    load_recipients_or_csv,
    validate_recipient_columns,
    generate_image_html,
)

DELAY_SEND = 10
DELAY_DRAFT = 1


def run_automailer(
    mode,
    recipients_path,
    exclusion_path,
    msg_template_path,
    progress_update,
    logger,
    embedded_images,
    real_attachments,
    pause_event,
    cancel_event,
    finish_callback,
    send_account_name,
    backend_type,
    smtp_host,
    smtp_port,
    smtp_user,
    smtp_pass,
    closing_statements,
):

    use_outlook = backend_type != "SMTP"
    if use_outlook:
        pythoncom.CoInitialize()
        backend = OutlookBackend(send_account_name)
    else:
        backend = SmtpBackend(smtp_host, int(smtp_port or 0), smtp_user, smtp_pass)

    try:
        recipients = load_recipients_or_csv(recipients_path, visible_only=True)
        validate_recipient_columns(recipients)
    except ValueError as e:
        messagebox.showerror("檔案錯誤", str(e))
        logger(f"收件人清單錯誤: {e}")
        if finish_callback:
            finish_callback(None, 0)
        return
    exclusion_emails = []
    if exclusion_path and os.path.exists(exclusion_path):
        try:
            exclusion_df = load_recipients_or_csv(exclusion_path)
            exclusion_emails = exclusion_df["Email"].tolist()
        except Exception as e:
            logger(f"排除清單讀取失敗: {e}")
    filtered = recipients[~recipients["Email"].isin(exclusion_emails)]

    msg = extract_msg.Message(msg_template_path)
    subject = msg.subject
    try:
        raw_html_body = msg.htmlBody
    except UnicodeDecodeError as e:
        logger(f"HTML 解析失敗: {e}")
        raw_html_body = None
    html_body = (
        raw_html_body.decode("utf-8", errors="ignore")
        if isinstance(raw_html_body, bytes)
        else (raw_html_body or "")
    )

    cid_list = list(embedded_images)
    image_html_all = generate_image_html(cid_list)

    total = len(filtered)

    """
    新增參數 cancel_event。每次迴圈開始前或 pause 時，都要檢查 cancel_event
    是否已被設置。設置就直接結束整個流程。
    新增一個參數 pause_event：threading.Event 物件。
    在每次實際要發送/存稿之前，都先呼叫 pause_event.wait()。
    當 pause_event 被 clear 時，wait() 會阻塞；被 set 時繼續執行。
    """
    last_index = None
    for i, row in filtered.iterrows():
        last_index = i
        # 若使用者按了「取消」，就直接跳出
        if cancel_event.is_set():
            logger("❌ 停止寄送，使用者已取消")
            break

        # 暫停處理：若 pause_event 沒被 set，就持續小睡並每次檢查 cancel_event
        while not pause_event.is_set():
            if cancel_event.is_set():
                logger("❌ 停止寄送，使用者已取消")
                break
            time.sleep(0.1)
        if cancel_event.is_set():
            break

        try:
            recipient = row["Email"]
            salutation = row["Salutation"]
            statement = random.choice(closing_statements)

            body = html_body.replace("[salutation]", salutation).replace(
                "[statement]", statement
            )

            def repl(match):
                idx = match.group(1)
                if idx == "":
                    return image_html_all
                try:
                    index = int(idx) - 1  # 讓 [image1] 代表第一張圖
                    if index < 0:
                        raise IndexError
                    cid = cid_list[index]
                    return generate_image_html([cid])
                except (ValueError, IndexError):
                    logger(f"⚠️ 無效的圖片佔位符：[image{idx}] → 找不到對應圖片")
                    return ""

            body = re.sub(r"\[image(\d*)\]", repl, body)

            backend.send(
                mode,
                recipient,
                subject,
                body,
                embedded_images,
                real_attachments,
            )
            logger(f"✉ 已處理：{recipient} / {salutation} / {statement}")
            progress_update(i, total, recipient)
            time.sleep(DELAY_SEND if mode == "send" else DELAY_DRAFT)
        except Exception as e:
            logger(f"❌ 寄送失敗：{recipient} - {e}")
            progress_update(i, total, f"{recipient} ❌")

    logger("✅ 所有郵件處理完成")

    if finish_callback:
        finish_callback(last_index, total)

    if use_outlook:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    from gui import GUI

    root = Tk()
    GUI(root)
    root.mainloop()
