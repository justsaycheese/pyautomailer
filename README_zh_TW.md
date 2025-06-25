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

使用以下指令安裝所需套件：

```bash
pip install -r requirements.txt
```

## 準備資料
### 收件者名單
建立 Excel（`.xlsx`/`.xls`）或 CSV 檔，至少包含兩欄：

- **Email** ─ 收件者地址
- **Salutation** ─ 在信件中使用的稱呼

在 Excel 中隱藏的列會被忽略。

### 排除名單 (可選)
可選的 Excel/CSV 檔，需含有 `Email` 欄位；會自動排除其中列出的地址。

### 郵件範本
使用 Outlook 的 `.msg` 檔作為郵件範本。HTML 內可以使用下列占位符，寄信時會自動替換：

- `[salutation]` ─ 以名單中的稱呼取代
- `[statement]` ─ 隨機結尾語
- `[image]` ─ 內嵌所有選擇圖片（若有選擇圖片）
- `[image1]`, `[image2]`, ...：插入指定編號圖片

> 若使用 `[image1]`, `[image2]`... 請確認有對應張數的圖片，否則會出現「錯誤佔位符」提示於日誌中。

> 若無選擇圖片會將其取代為空字元。


若範本中的 RTF 內容包含無法解碼的位元組，程式會自動忽略該部分以避免錯誤。

## 執行方式
在終端機輸入：

```bash
python automailer.py
```
或直接下載並執行已發佈的`.exe`檔

### 操作介面說明
- 可選 Outlook 或 SMTP 模式寄信
- 支援圖片及附件資料夾或多檔案載入
- 寄送模式可選「寄出」或「儲存草稿」

### 設定儲存
使用者設定（寄件帳號、檔案路徑等）會儲存於 `settings.json`，可透過按鈕儲存，下次開啟自動載入。

### 平台限制
- Outlook 模式僅限 Windows 且需安裝 Outlook。
- SMTP 模式則無平台限制，只需能連上郵件伺服器。


## 範例文本
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
