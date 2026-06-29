# PrinterSelector — 黑白 / 彩色列印選擇器

## 功能
- 選擇檔案後顯示預覽（PDF、圖片、Word/Excel/PPT）
- 一鍵送至黑白或彩色印表機
- 印表機名稱可在設定中選擇並儲存

## 支援格式
| 格式 | 預覽 | 備註 |
|------|------|------|
| PDF | ✅ | 內建 |
| 圖片（JPG/PNG/BMP/GIF/TIFF）| ✅ | 內建 |
| Word（.doc/.docx）| ✅ | 需要安裝 Microsoft Word |
| Excel（.xls/.xlsx）| ✅ | 需要安裝 Microsoft Excel |
| PowerPoint（.ppt/.pptx）| ✅ | 需要安裝 Microsoft PowerPoint |

## 取得 EXE（GitHub Actions）

### 步驟一：建立 GitHub Repository
1. 登入 [github.com](https://github.com)
2. 右上角 `+` → `New repository`
3. 名稱填 `printer-selector`，Visibility 選 **Public** 或 **Private** 皆可
4. 點 `Create repository`

### 步驟二：上傳程式碼
在你的 Windows 電腦上（需要安裝 [Git](https://git-scm.com/download/win)）：

```bat
cd 你存放這個資料夾的位置
git init
git add .
git commit -m "init"
git branch -M main
git remote add origin https://github.com/你的帳號/printer-selector.git
git push -u origin main
```

### 步驟三：等待自動打包
- push 完成後，進到 GitHub repo 頁面
- 點上方 `Actions` 標籤
- 等 `Build Windows EXE` 跑完（約 3–5 分鐘）

### 步驟四：下載 EXE
- Actions 頁面點進該次 workflow run
- 最下方 `Artifacts` → 點 `PrinterSelector-Windows` 下載
- 解壓縮後執行 `PrinterSelector.exe`

## 設定檔位置
`%APPDATA%\PrinterSelector\config.json`

## 系統需求
- Windows 10 / 11
- 預覽 Office 格式需要安裝對應的 Microsoft Office
