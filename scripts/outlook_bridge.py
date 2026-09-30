"""
Outlook M365 Bridge Utility Script
專為企業環境打造的 Outlook M365 整合工具
支援：讀取郵件、關鍵字搜尋、寄送郵件、行事曆查詢與建立、以及背景程序防呆清理
"""

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# Ensure UTF-8 output on Windows terminal
sys.stdout.reconfigure(encoding="utf-8")

ONEDRIVE_DEFAULT = Path(r"C:\Users\jeffho\OneDrive - ASSA ABLOY Group")
EMAIL_FILE = ONEDRIVE_DEFAULT / "latest_emails.json"


def run_ps(script: str) -> str:
    """執行 PowerShell 腳本並回傳 stdout"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True,
        text=True,
        timeout=30,
        encoding="utf-8",
        errors="replace"
    )
    if result.returncode != 0 and result.stderr:
        raise RuntimeError(result.stderr.strip())
    return result.stdout.strip()


def get_sender_display(email_item: dict) -> str:
    from_field = email_item.get("from")
    if isinstance(from_field, dict):
        addr_info = from_field.get("emailAddress", {})
        name = addr_info.get("name", "")
        addr = addr_info.get("address", "")
        return f"{name} <{addr}>" if name else addr
    elif isinstance(from_field, str):
        return from_field
    return "未知寄件人"


def clean_zombies():
    """清理背景殘留的 OUTLOOK.EXE 殭屍程序，防止 MAPI 1.0 [000004C2] 鎖定錯誤"""
    ps = """
    $procs = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
    if ($procs) {
        $count = $procs.Count
        Stop-Process -Name 'OUTLOOK' -Force -ErrorAction SilentlyContinue
        Write-Host "已清理 $count 個背景殘留的 OUTLOOK 程序。"
    } else {
        Write-Host "目前無殘留的 OUTLOOK 背景程序。"
    }
    """
    print(run_ps(ps))


def check_status():
    """檢查整體連線狀態與檔案同步時間"""
    print("=" * 60)
    print(" Outlook M365 Bridge 狀態檢查")
    print("=" * 60)
    
    # 檢查 OneDrive 檔案
    if EMAIL_FILE.exists():
        stat = EMAIL_FILE.stat()
        mtime = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        size_kb = round(stat.st_size / 1024, 1)
        print(f"✅ OneDrive 郵件快取檔：已就緒")
        print(f"   路徑：{EMAIL_FILE}")
        print(f"   最後更新：{mtime}")
        print(f"   檔案大小：{size_kb} KB")
        try:
            with open(EMAIL_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            print(f"   包含郵件數：{len(data)} 封")
        except Exception as e:
            print(f"   ⚠️ 解析失敗：{e}")
    else:
        print(f"⚠️ OneDrive 郵件快取檔尚未生成（等待 Power Automate 觸發）：{EMAIL_FILE}")

    # 檢查本機 Outlook 狀態
    ps_check = """
    $procs = Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue
    if ($procs) {
        Write-Host "Outlook 處理程序數: $($procs.Count)"
    } else {
        Write-Host "Outlook 目前未在執行中。"
    }
    """
    print(f"\n💻 本機程序狀態：\n   {run_ps(ps_check)}")
    print("=" * 60)


def read_emails(count: int = 5, keyword: str = "", unread_only: bool = False):
    """讀取 OneDrive 同步的郵件"""
    if not EMAIL_FILE.exists():
        print(f"❌ 找不到郵件快取檔案：{EMAIL_FILE}")
        print("請確認 Power Automate「Antigravity-Outlook-Bridge」流程已啟用。")
        return

    stat = EMAIL_FILE.stat()
    last_updated = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")

    with open(EMAIL_FILE, "r", encoding="utf-8") as f:
        emails = json.load(f)

    if not emails:
        print("收件匣目前無郵件。")
        return

    # 未讀過濾
    if unread_only:
        emails = [e for e in emails if not e.get("isRead", True)]

    # 關鍵字過濾
    if keyword:
        kw = keyword.lower()
        emails = [
            e for e in emails
            if kw in (e.get("subject", "") or "").lower()
            or kw in str(e.get("from", "")).lower()
            or kw in (e.get("bodyPreview", "") or "").lower()
        ]
        if not emails:
            print(f"找不到包含關鍵字「{keyword}」的快取郵件。")
            return

    emails = emails[:count]
    print(f"📬 最新 Outlook 郵件清單（快取更新時間：{last_updated}）")
    print("─" * 60)
    for i, e in enumerate(emails, 1):
        subj = e.get("subject", "（無主旨）")
        sender = get_sender_display(e)
        received = e.get("receivedDateTime", "")
        is_read = e.get("isRead", True)
        has_attach = " 📎" if e.get("hasAttachments") else ""
        preview = (e.get("bodyPreview") or "").replace("\r", " ").replace("\n", " ").strip()

        unread_mark = "🔵 [未讀] " if not is_read else ""
        print(f"\n[{i}] {unread_mark}{received[:16].replace('T', ' ')}{has_attach}")
        print(f"    寄件人：{sender}")
        print(f"    主旨：{subj}")
        print(f"    摘要：{preview[:130]}...")


def search_inbox_live(keyword: str, count: int = 15):
    """透過 PowerShell DASL 查詢即時掃描 Outlook 收件匣全量主旨"""
    ps = f"""
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace('MAPI')
    $account = $ns.Folders.Item(2)
    $inbox = $account.Folders.Item('收件匣')
    $filter = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%{keyword}%'"
    $items = $inbox.Items.Restrict($filter)
    $items.Sort('[ReceivedTime]', $true)
    $total = $items.Count
    Write-Host "TOTAL:$total"
    for ($i = 1; $i -le [Math]::Min({count}, $total); $i++) {{
        $m = $items.Item($i)
        $t = $m.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss')
        $s = $m.Subject
        $sz = [Math]::Round($m.Size / 1024, 1)
        $u = if ($m.UnRead) {{ "未讀" }} else {{ "已讀" }}
        Write-Host "ITEM|$i|$t|$u|${{sz}}KB|$s"
    }}
    """
    try:
        raw = run_ps(ps)
        lines = [line.strip() for line in raw.splitlines() if line.strip()]
        total = 0
        items = []
        for line in lines:
            if line.startswith("TOTAL:"):
                total = int(line.split(":", 1)[1])
            elif line.startswith("ITEM|"):
                parts = line.split("|", 5)
                items.append(parts[1:])

        print(f"🔍 全量收件匣搜尋「{keyword}」：共找到 {total} 封郵件")
        print("─" * 70)
        for it in items:
            idx, t, u, sz, subj = it
            u_tag = "🔵 " if u == "未讀" else "   "
            print(f"[{idx:>2}] {u_tag}{t} ({sz}) | {subj}")
    except Exception as e:
        print(f"搜尋失敗：{e}")


def send_email(to: str, subject: str, body: str, cc: str = ""):
    """寄送郵件"""
    sub_esc = subject.replace("'", "''")
    body_esc = body.replace("'", "''")
    cc_part = f"$mail.CC = '{cc}'" if cc else ""
    ps = f"""
    $outlook = New-Object -ComObject Outlook.Application
    $mail = $outlook.CreateItem(0)
    $mail.To = '{to}'
    $mail.Subject = '{sub_esc}'
    $mail.Body = '{body_esc}'
    {cc_part}
    $mail.Send()
    Write-Host 'SENT_SUCCESS'
    """
    try:
        out = run_ps(ps)
        if "SENT_SUCCESS" in out:
            print(f"✅ 郵件已成功排入寄送：\n   收件人：{to}\n   主旨：{subject}")
        else:
            print(f"⚠️ 寄送回應：{out}")
    except Exception as e:
        print(f"❌ 寄送失敗：{e}")


def main():
    parser = argparse.ArgumentParser(description="Outlook M365 Bridge 整合工具")
    subparsers = parser.add_subparsers(dest="command", help="子指令")

    # read
    p_read = subparsers.add_parser("read", help="讀取最新郵件")
    p_read.add_argument("-n", "--count", type=int, default=5, help="讀取數量（預設 5）")
    p_read.add_argument("-k", "--keyword", type=str, default="", help="過濾關鍵字")
    p_read.add_argument("-u", "--unread", action="store_true", help="只顯示未讀")

    # search
    p_search = subparsers.add_parser("search", help="即時搜尋 Outlook 收件匣主旨")
    p_search.add_argument("keyword", type=str, help="搜尋關鍵字")
    p_search.add_argument("-n", "--count", type=int, default=15, help="最多顯示幾筆")

    # send
    p_send = subparsers.add_parser("send", help="寄送電子郵件")
    p_send.add_argument("--to", required=True, help="收件人 Email")
    p_send.add_argument("--subject", required=True, help="郵件主旨")
    p_send.add_argument("--body", required=True, help="郵件內容")
    p_send.add_argument("--cc", default="", help="副本 Email")

    # status
    subparsers.add_parser("status", help="檢查橋接同步狀態")

    # clean-zombies
    subparsers.add_parser("clean-zombies", help="清理殘留的 OUTLOOK.EXE 避免鎖定")

    args = parser.parse_args()

    if args.command == "read":
        read_emails(count=args.count, keyword=args.keyword, unread_only=args.unread)
    elif args.command == "search":
        search_inbox_live(args.keyword, count=args.count)
    elif args.command == "send":
        send_email(to=args.to, subject=args.subject, body=args.body, cc=args.cc)
    elif args.command == "status":
        check_status()
    elif args.command == "clean-zombies":
        clean_zombies()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
