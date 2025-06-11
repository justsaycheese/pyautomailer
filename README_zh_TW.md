# Automailer

此專案提供一個 GUI 工具，可透過 Outlook 或任意 SMTP 伺服器批次寄送電子郵件。

## 系統需求
- Python 3.10 以上
- 套件：
  - `extract_msg`
  - `pandas`
  - `openpyxl`
  - `pywin32`（僅 Outlook 模式需要）
  - `tkinter`（Windows 版 Python 內建）

使用以下指令安裝所需套件（Outlook 模式才需安裝 pywin32）：

```bash
pip install extract_msg pandas openpyxl pywin32
```

## 準備資料
### 收件者名單
建立 Excel（`.xlsx`/`.xls`）或 CSV 檔，至少包含兩欄：

- **Email** ─ 收件者地址
- **Salutation** ─ 在信件中使用的稱呼

在 Excel 中隱藏的列會被忽略。

### 排除名單
可選的 Excel/CSV 檔，需含有 `Email` 欄位；其中列出的地址會自動排除。

### 郵件範本
使用 Outlook 的 `.msg` 檔作為郵件範本。HTML 內可以使用下列占位符，寄信時會自動替換：

- `[salutation]` ─ 以名單中的稱呼取代
- `[statement]` ─ 隨機結尾語
- `[image]` ─ 內嵌圖片的 HTML（若有選擇圖片）

若範本中的 RTF 內容包含無法解碼的位元組，程式會自動忽略該部分以避免錯誤。

## 執行方式
在終端機輸入：

```bash
python automailer_verZ.py
```

啟動後在介面中可選擇 **Outlook** 或 **SMTP** 模式。
若使用 SMTP，請輸入主機、連接埠與登入資訊，再選擇收件者名單及郵件範本，按下 **Start** 即可寄送或存檔。

### 跨平台說明
Outlook 模式僅能在安裝 Outlook 的 Windows 系統使用。
SMTP 模式則可在任何平台運作，只要能連上郵件伺服器。
