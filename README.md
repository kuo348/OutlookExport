# Outlook 郵件匯出工具

將 Outlook（傳統桌面版）的郵件依原本的資料夾結構匯出到電腦，每封郵件存成獨立資料夾，包含 `.msg`、內文、郵件資訊與附件，並產生 CSV 索引。

## 需求
- Windows + 傳統版 Outlook
- Python 3
- `pip install -r requirements.txt`

## 用法
```powershell
python outlook_export.py --list                        # 列出所有 Outlook 資料夾
python outlook_export.py -o D:\郵件備份                 # 全部匯出
python outlook_export.py -f "收件匣/專案A" -o D:\郵件備份 # 只匯出指定資料夾（含子資料夾）
python outlook_export.py --since 2026-01-01            # 只匯出某日期之後
python outlook_export.py --flat                        # 每封只存一個 .msg
python outlook_export.py --overwrite                   # 重新匯出已存在的郵件
```

## 輸出結構
```
輸出資料夾\
├─ 郵件索引_日期時間.csv
└─ 帳號\收件匣\子資料夾\
   └─ 2026-01-05_093012_主旨\
      ├─ 郵件.msg
      ├─ 內文.txt / 內文.html
      ├─ 郵件資訊.txt
      └─ 附件\
```
