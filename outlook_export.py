# -*- coding: utf-8 -*-
r"""
Outlook 郵件匯出工具
將 Outlook 的郵件資料夾結構轉成電腦上的資料夾，每封郵件存成檔案（含附件）。

需求：Windows + 傳統版 Outlook（桌面版）+ pip install pywin32

用法範例：
    python outlook_export.py                         # 匯出所有帳號、所有資料夾到 .\OutlookExport_輸出
    python outlook_export.py -o D:\郵件備份           # 指定輸出位置
    python outlook_export.py -f "收件匣"              # 只匯出收件匣（含子資料夾）
    python outlook_export.py -f "收件匣/專案A"        # 只匯出指定子資料夾
    python outlook_export.py --since 2026-01-01      # 只匯出某日期之後的郵件
    python outlook_export.py --flat                  # 每封郵件只存一個 .msg 檔，不另開子資料夾
    python outlook_export.py --list                  # 列出所有 Outlook 資料夾
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys

try:
    import win32com.client
    import pywintypes
except ImportError:
    sys.exit("缺少 pywin32，請先執行：python -m pip install pywin32")

# Outlook 常數
OL_MSG_UNICODE = 9      # .msg（Unicode）
OL_MAIL_ITEM = 43       # MailItem class
OL_FORMAT_HTML = 2

INVALID_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')
MAX_NAME = 80           # 單一名稱最大長度，避免超過 Windows 路徑長度限制


def safe_name(name, max_len=MAX_NAME):
    """把字串轉成合法的 Windows 檔名/資料夾名稱。"""
    name = INVALID_CHARS.sub("_", str(name or "")).strip().rstrip(". ")
    name = re.sub(r"\s+", " ", name)
    return (name[:max_len].rstrip(". ") or "(無主旨)")


def unique_path(path):
    """若路徑已存在，自動加上 (2)、(3)…"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{base} ({i}){ext}"):
        i += 1
    return f"{base} ({i}){ext}"


def to_datetime(value):
    try:
        return dt.datetime(value.year, value.month, value.day,
                           value.hour, value.minute, value.second)
    except Exception:
        return None


def long_path(path):
    """加上 \\\\?\\ 前綴以支援超過 260 字元的路徑。"""
    path = os.path.abspath(path)
    return path if path.startswith("\\\\?\\") else "\\\\?\\" + path


class Exporter:
    def __init__(self, out_dir, since=None, flat=False, skip_existing=True):
        self.out_dir = os.path.abspath(out_dir)
        self.since = since
        self.flat = flat
        self.skip_existing = skip_existing
        self.count = 0
        self.errors = 0
        self.index_rows = []
        outlook = win32com.client.Dispatch("Outlook.Application")
        self.ns = outlook.GetNamespace("MAPI")

    # ---------- 資料夾 ----------
    def roots(self):
        return [self.ns.Folders.Item(i + 1) for i in range(self.ns.Folders.Count)]

    def list_folders(self, folder=None, depth=0):
        folders = [folder] if folder else self.roots()
        for f in folders:
            try:
                n = f.Items.Count
            except Exception:
                n = "?"
            print("  " * depth + f"{f.Name}  ({n})")
            for i in range(f.Folders.Count):
                self.list_folders(f.Folders.Item(i + 1), depth + 1)

    def find_folder(self, path):
        """依 "帳號/收件匣/子資料夾" 或 "收件匣/子資料夾" 尋找資料夾。"""
        parts = [p for p in re.split(r"[\\/]", path) if p]
        candidates = self.roots()
        # 若第一段是帳號名稱，就從該帳號開始找
        for root in candidates:
            if root.Name.lower() == parts[0].lower():
                return self._walk(root, parts[1:]), root.Name
        for root in candidates:
            f = self._walk(root, parts)
            if f is not None:
                return f, root.Name
        return None, None

    @staticmethod
    def _walk(folder, parts):
        for p in parts:
            nxt = None
            for i in range(folder.Folders.Count):
                sub = folder.Folders.Item(i + 1)
                if sub.Name.lower() == p.lower():
                    nxt = sub
                    break
            if nxt is None:
                return None
            folder = nxt
        return folder

    # ---------- 匯出 ----------
    def export_folder(self, folder, disk_dir):
        os.makedirs(long_path(disk_dir), exist_ok=True)
        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass
        total = items.Count
        print(f"\n📁 {folder.FolderPath}  ({total} 個項目)")

        for i in range(1, total + 1):
            try:
                item = items.Item(i)
            except Exception:
                continue
            try:
                self.export_item(item, disk_dir)
            except pywintypes.com_error as e:
                self.errors += 1
                print(f"   ⚠ 略過一封郵件（COM 錯誤）：{e}")
            except Exception as e:
                self.errors += 1
                print(f"   ⚠ 略過一封郵件：{e}")

        for i in range(folder.Folders.Count):
            sub = folder.Folders.Item(i + 1)
            self.export_folder(sub, os.path.join(disk_dir, safe_name(sub.Name)))

    def export_item(self, item, disk_dir):
        when = to_datetime(getattr(item, "ReceivedTime", None) or getattr(item, "CreationTime", None))
        if self.since and when and when < self.since:
            return

        subject = getattr(item, "Subject", "") or "(無主旨)"
        stamp = when.strftime("%Y-%m-%d_%H%M%S") if when else "未知日期"
        base = f"{stamp}_{safe_name(subject, 60)}"

        if self.flat:
            msg_path = os.path.join(disk_dir, base + ".msg")
            if self.skip_existing and os.path.exists(long_path(msg_path)):
                return
            item.SaveAs(msg_path, OL_MSG_UNICODE)
            mail_dir, attach_names = disk_dir, []
        else:
            mail_dir = os.path.join(disk_dir, base)
            if self.skip_existing and os.path.exists(long_path(mail_dir)):
                return
            os.makedirs(long_path(mail_dir), exist_ok=True)
            msg_path = os.path.join(mail_dir, "郵件.msg")
            item.SaveAs(msg_path, OL_MSG_UNICODE)
            self._save_body(item, mail_dir)
            self._save_info(item, mail_dir, when)
            attach_names = self._save_attachments(item, os.path.join(mail_dir, "附件"))

        self.count += 1
        self.index_rows.append({
            "日期": when.strftime("%Y-%m-%d %H:%M:%S") if when else "",
            "寄件者": getattr(item, "SenderName", ""),
            "主旨": subject,
            "Outlook資料夾": item.Parent.FolderPath,
            "附件": "; ".join(attach_names),
            "存放位置": os.path.relpath(msg_path, self.out_dir),
        })
        if self.count % 50 == 0:
            print(f"   …已匯出 {self.count} 封")

    def _save_body(self, item, mail_dir):
        try:
            if getattr(item, "BodyFormat", None) == OL_FORMAT_HTML and item.HTMLBody:
                with open(long_path(os.path.join(mail_dir, "內文.html")), "w", encoding="utf-8") as f:
                    html = item.HTMLBody
                    if "charset" not in html[:1000].lower():
                        html = '<meta charset="utf-8">\n' + html
                    f.write(html)
            with open(long_path(os.path.join(mail_dir, "內文.txt")), "w", encoding="utf-8") as f:
                f.write(getattr(item, "Body", "") or "")
        except Exception as e:
            print(f"   ⚠ 內文儲存失敗：{e}")

    def _save_info(self, item, mail_dir, when):
        def g(attr):
            try:
                return getattr(item, attr, "") or ""
            except Exception:
                return ""
        lines = [
            f"主旨：{g('Subject')}",
            f"寄件者：{g('SenderName')} <{g('SenderEmailAddress')}>",
            f"收件者：{g('To')}",
            f"副本：{g('CC')}",
            f"日期：{when or ''}",
            f"Outlook 資料夾：{item.Parent.FolderPath}",
        ]
        with open(long_path(os.path.join(mail_dir, "郵件資訊.txt")), "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")

    def _save_attachments(self, item, att_dir):
        names = []
        try:
            atts = item.Attachments
            count = atts.Count
        except Exception:
            return names
        if count == 0:
            return names
        os.makedirs(long_path(att_dir), exist_ok=True)
        for i in range(1, count + 1):
            att = atts.Item(i)
            try:
                fname = att.FileName or f"附件{i}"
            except Exception:
                fname = f"附件{i}"
            root, ext = os.path.splitext(safe_name(fname, 120))
            path = unique_path(os.path.join(att_dir, root[:80] + ext))
            try:
                att.SaveAsFile(path)
                names.append(os.path.basename(path))
            except Exception as e:
                print(f"   ⚠ 附件「{fname}」儲存失敗：{e}")
        if not os.listdir(long_path(att_dir)):
            os.rmdir(long_path(att_dir))
        return names

    def write_index(self):
        if not self.index_rows:
            return None
        path = os.path.join(self.out_dir, f"郵件索引_{dt.datetime.now():%Y%m%d_%H%M%S}.csv")
        with open(path, "w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig：Excel 開啟不亂碼
            w = csv.DictWriter(f, fieldnames=list(self.index_rows[0].keys()))
            w.writeheader()
            w.writerows(self.index_rows)
        return path


def main():
    ap = argparse.ArgumentParser(description="將 Outlook 郵件匯出成資料夾與檔案")
    ap.add_argument("-o", "--output", default="OutlookExport_輸出", help="輸出資料夾（預設：.\\OutlookExport_輸出）")
    ap.add_argument("-f", "--folder", help='只匯出指定的 Outlook 資料夾，例如 "收件匣" 或 "帳號/收件匣/專案A"')
    ap.add_argument("--since", help="只匯出此日期之後的郵件，格式 YYYY-MM-DD")
    ap.add_argument("--flat", action="store_true", help="每封郵件只存一個 .msg 檔，不建立個別子資料夾")
    ap.add_argument("--overwrite", action="store_true", help="重新匯出已存在的郵件（預設會略過，方便續傳）")
    ap.add_argument("--list", action="store_true", help="列出所有 Outlook 資料夾後結束")
    args = ap.parse_args()

    since = dt.datetime.strptime(args.since, "%Y-%m-%d") if args.since else None
    ex = Exporter(args.output, since=since, flat=args.flat, skip_existing=not args.overwrite)

    if args.list:
        ex.list_folders()
        return

    os.makedirs(ex.out_dir, exist_ok=True)
    start = dt.datetime.now()

    if args.folder:
        folder, account = ex.find_folder(args.folder)
        if folder is None:
            sys.exit(f"找不到 Outlook 資料夾：{args.folder}（可用 --list 查看所有資料夾）")
        ex.export_folder(folder, os.path.join(ex.out_dir, safe_name(account), safe_name(folder.Name)))
    else:
        for root in ex.roots():
            ex.export_folder(root, os.path.join(ex.out_dir, safe_name(root.Name)))

    index = ex.write_index()
    secs = (dt.datetime.now() - start).total_seconds()
    print(f"\n✅ 完成！共匯出 {ex.count} 封郵件，失敗 {ex.errors} 封，耗時 {secs:.0f} 秒")
    print(f"   輸出位置：{ex.out_dir}")
    if index:
        print(f"   郵件索引：{index}")


if __name__ == "__main__":
    main()
