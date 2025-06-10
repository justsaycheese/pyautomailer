# Automailer

此專案提供一個 GUI 工具，可透過 Outlook 批次寄送電子郵件。

## 系統需求
- Python 3.10 以上
- 套件：
  - `extract_msg`
  - `pandas`
  - `openpyxl`
  - `pywin32`（提供 `win32com.client`）
  - `tkinter`（Windows 版 Python 內建）

使用以下指令安裝所需套件：

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

## 執行方式
在終端機輸入：

```bash
python automailer_verZ.py
```

啟動後在介面中選擇收件者名單、（可選）排除名單及郵件範本，亦可加入要內嵌的圖片或附檔。最後按下 **Start** 便會利用 Outlook 進行寄送或建立草稿。
