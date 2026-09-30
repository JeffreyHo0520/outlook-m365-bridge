"""
Outlook Bridge MCP Server for Antigravity
結合 Power Automate 雲端同步與 PowerShell 本機操作：
- 讀取郵件與搜尋：從 OneDrive 自動同步的 latest_emails.json 讀取（繞過本機 COM 限制，具備完整內文、寄件人與附件資訊）
- 發送郵件與行事曆：透過本機 PowerShell COM 自動化執行
"""

import asyncio
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

app = Server("outlook-bridge")

ONEDRIVE_PATH = Path(r"C:\Users\jeffho\OneDrive - ASSA ABLOY Group")
EMAIL_FILE = ONEDRIVE_PATH / "latest_emails.json"


def run_ps(script: str) -> str:
    """執行本機 PowerShell 腳本"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, timeout=30,
        encoding="utf-8", errors="replace"
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


# ──────────────────────────────────────────
# 工具宣告
# ──────────────────────────────────────────

@app.list_tools()
async def list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="get_inbox_emails",
            description="讀取 Outlook 收件匣最新郵件（包含主旨、寄件人、時間、摘要與附件狀態）",
            inputSchema={
                "type": "object",
                "properties": {
                    "count": {
                        "type": "integer",
                        "description": "讀取筆數（預設 5）",
                        "default": 5
                    },
                    "unread_only": {
                        "type": "boolean",
                        "description": "是否僅列出未讀信件",
                        "default": False
                    }
                }
            }
        ),
        types.Tool(
            name="search_emails",
            description="搜尋收件匣郵件（關鍵字可比對主旨、寄件人或內文預覽）",
            inputSchema={
                "type": "object",
                "properties": {
                    "keyword": {
                        "type": "string",
                        "description": "搜尋關鍵字"
                    },
                    "count": {
                        "type": "integer",
                        "description": "最多回傳幾筆",
                        "default": 10
                    }
                },
                "required": ["keyword"]
            }
        ),
        types.Tool(
            name="send_email",
            description="透過本機 Outlook 寄送郵件",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "收件人 Email"},
                    "subject": {"type": "string", "description": "主旨"},
                    "body": {"type": "string", "description": "內文"},
                    "cc": {"type": "string", "description": "副本（可選）", "default": ""}
                },
                "required": ["to", "subject", "body"]
            }
        ),
        types.Tool(
            name="get_calendar_events",
            description="讀取 Outlook 行事曆的近期行程",
            inputSchema={
                "type": "object",
                "properties": {
                    "days_ahead": {
                        "type": "integer",
                        "description": "讀取未來天數（預設 7）",
                        "default": 7
                    }
                }
            }
        ),
        types.Tool(
            name="create_calendar_event",
            description="在 Outlook 行事曆新增行程",
            inputSchema={
                "type": "object",
                "properties": {
                    "subject": {"type": "string", "description": "行程主旨"},
                    "start_datetime": {"type": "string", "description": "開始時間（YYYY-MM-DD HH:MM）"},
                    "end_datetime": {"type": "string", "description": "結束時間（YYYY-MM-DD HH:MM）"},
                    "location": {"type": "string", "description": "地點（可選）", "default": ""},
                    "body": {"type": "string", "description": "備註（可選）", "default": ""}
                },
                "required": ["subject", "start_datetime", "end_datetime"]
            }
        )
    ]


# ──────────────────────────────────────────
# 工具實作
# ──────────────────────────────────────────

@app.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:

    if name in ("get_inbox_emails", "search_emails"):
        if not EMAIL_FILE.exists():
            return [types.TextContent(type="text", text="❌ 尚未找到郵件同步檔案，請確認 Power Automate 流程已執行。")]

        stat = EMAIL_FILE.stat()
        last_updated = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")

        with open(EMAIL_FILE, "r", encoding="utf-8") as f:
            emails = json.load(f)

        if not emails:
            return [types.TextContent(type="text", text="收件匣目前無郵件。")]

        if name == "search_emails":
            kw = arguments.get("keyword", "").lower()
            emails = [
                e for e in emails
                if kw in (e.get("subject", "") or "").lower()
                or kw in str(e.get("from", "")).lower()
                or kw in (e.get("bodyPreview", "") or "").lower()
            ]
            if not emails:
                return [types.TextContent(type="text", text=f"找不到包含關鍵字「{arguments.get('keyword')}」的郵件。")]

        elif name == "get_inbox_emails":
            if arguments.get("unread_only", False):
                emails = [e for e in emails if not e.get("isRead", True)]

        count = arguments.get("count", 5)
        emails = emails[:count]

        lines = [f"📬 最新 Outlook 郵件（更新時間：{last_updated}）\n{'─'*55}"]
        for i, email in enumerate(emails, 1):
            subject = email.get("subject", "（無主旨）")
            sender = get_sender_display(email)
            received = email.get("receivedDateTime", "")
            is_read = email.get("isRead", True)
            has_attach = " 📎" if email.get("hasAttachments") else ""
            preview = (email.get("bodyPreview") or "").replace("\r", " ").replace("\n", " ").strip()

            unread_mark = "🔵 [未讀] " if not is_read else ""
            lines.append(
                f"\n[{i}] {unread_mark}{received[:16].replace('T', ' ')}{has_attach}\n"
                f"    寄件人：{sender}\n"
                f"    主旨：{subject}\n"
                f"    摘要：{preview[:140]}..."
            )
        return [types.TextContent(type="text", text="\n".join(lines))]

    elif name == "send_email":
        to = arguments["to"]
        subject = arguments["subject"].replace("'", "''")
        body = arguments["body"].replace("'", "''")
        cc = arguments.get("cc", "")
        cc_cmd = f"$mail.CC = '{cc}'" if cc else ""
        ps = f"""
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = '{to}'
$mail.Subject = '{subject}'
$mail.Body = '{body}'
{cc_cmd}
$mail.Send()
Write-Host 'SENT'
"""
        try:
            run_ps(ps)
            return [types.TextContent(type="text", text=f"✅ 郵件已成功寄出給 {to}！")]
        except Exception as ex:
            return [types.TextContent(type="text", text=f"寄送失敗：{ex}")]

    elif name == "get_calendar_events":
        days = arguments.get("days_ahead", 7)
        ps = f"""
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('MAPI')
$cal = $ns.GetDefaultFolder(9)
$items = $cal.Items
$items.IncludeRecurrences = $true
$items.Sort('[Start]')
$now = Get-Date
$end = $now.AddDays({days})
$results = @()
foreach ($item in $items) {{
    try {{
        if ($item.Start -ge $now -and $item.Start -le $end) {{
            $results += [PSCustomObject]@{{
                Start    = $item.Start.ToString('yyyy-MM-dd HH:mm')
                End      = $item.End.ToString('HH:mm')
                Subject  = $item.Subject
                Location = $item.Location
            }}
        }}
    }} catch {{}}
    if ($results.Count -ge 20) {{ break }}
}}
$results | ConvertTo-Json -Compress
"""
        try:
            out = run_ps(ps)
            if not out or out == "null":
                return [types.TextContent(type="text", text=f"未來 {days} 天內無排定行程。")]
            events = json.loads(out)
            if isinstance(events, dict):
                events = [events]
            lines = [f"📅 未來 {days} 天行程（共 {len(events)} 筆）\n{'─'*55}"]
            for e in events:
                loc = f" @ {e['Location']}" if e.get("Location") else ""
                lines.append(f"\n• {e['Start']} ~ {e['End']}{loc}\n  {e['Subject']}")
            return [types.TextContent(type="text", text="\n".join(lines))]
        except Exception as ex:
            return [types.TextContent(type="text", text=f"讀取行事曆失敗：{ex}")]

    elif name == "create_calendar_event":
        subj = arguments["subject"].replace("'", "''")
        start = arguments["start_datetime"]
        end = arguments["end_datetime"]
        loc = arguments.get("location", "").replace("'", "''")
        body = arguments.get("body", "").replace("'", "''")
        ps = f"""
$outlook = New-Object -ComObject Outlook.Application
$appt = $outlook.CreateItem(1)
$appt.Subject = '{subj}'
$appt.Start = '{start}'
$appt.End = '{end}'
$appt.Location = '{loc}'
$appt.Body = '{body}'
$appt.Save()
Write-Host 'SAVED'
"""
        try:
            run_ps(ps)
            return [types.TextContent(type="text", text=f"✅ 行程已成功建立：{arguments['subject']}（{start} ~ {end}）")]
        except Exception as ex:
            return [types.TextContent(type="text", text=f"建立行程失敗：{ex}")]

    return [types.TextContent(type="text", text=f"未知工具：{name}")]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(main())
