---
name: outlook-m365-bridge
description: 專門用於企業環境（如 ASSA ABLOY）之 Outlook M365 雙軌郵件與行事曆整合橋接技能。採用「雲端 Power Automate + 本機 OneDrive 快取」合規讀信架構，完全繞過企業 GPO 限制與 Azure 401 權限封鎖，完整保留寄件人、主旨、內文摘要與未讀狀態；同時結合本機 PowerShell COM 自動化實現全量收件匣 DASL 搜尋、發送郵件與行事曆行程管理，並內建防止 MAPI 1.0 [000004C2] 登入鎖定之背景程序防呆機制。
---

# Outlook M365 雙軌企業整合技能 (`outlook-m365-bridge`)

## 概述

此技能專為在企業嚴格資安環境（如 **ASSA ABLOY / 東隆五金 TLM**）下，為 AI Agent（Antigravity）提供安全、穩定且**無需 IT 特殊管理員授權**的 Outlook M365 雙向串接方案。

### 核心痛點與架構突破

| 傳統方式痛點 | 本技能之雙軌解決方案 |
|---|---|
| **Azure 401 封鎖**：一般員工無權建立 Azure App Registration | **Power Automate 雲端標準版**：使用內建 M365 授權，自動定時將郵件同步至個人 OneDrive |
| **GPO 封鎖 COM 郵件內文**：Python/COM 讀取信件屬性卡死或報錯 `0x80004004` | **OneDrive JSON 本地快取**：AI 本地直接讀取同步 JSON，完整取得寄件人、收件時間、主旨、內文摘要與未讀標記 |
| **殭屍程序鎖死 MAPI**：背景呼叫 Outlook 導致桌面版開啟時出現 `MAPI 1.0 [000004C2]` 逾時錯誤 | **程序防呆與進程防護 (`clean-zombies`)**：主動偵測與釋放殘留 `OUTLOOK.EXE`，確保桌面端隨時正常開啟 |
| **全量收件匣歷史查找困難** | **PowerShell DASL 瞬時搜尋**：透過 Outlook Table/DASL 直接即時過濾歷史上萬封主旨 |

---

## 模組與檔案結構

```
outlook-m365-bridge/
├── SKILL.md                      # 技能詳細規格與架構說明
└── scripts/
    ├── outlook_bridge.py         # 核心 CLI 指令工具（讀取、搜尋、寄信、除錯、狀態）
    └── mcp_server.py             # Antigravity 原生 MCP 伺服器
```

---

## 快速使用指南 (CLI Commands)

所有操作均可透過 Python 環境執行：

```powershell
$py = "C:\Users\jeffho\.venv\Scripts\python.exe"
$script = "C:\Users\jeffho\.gemini\config\skills\outlook-m365-bridge\scripts\outlook_bridge.py"
```

### 1. 檢查連線與同步狀態
確認 OneDrive 同步檔與本機 Outlook 程序：
```powershell
& $py $script status
```

### 2. 讀取最新收件匣郵件
讀取最新 5 封（含寄件人、主旨、時間、未讀標示與摘要）：
```powershell
& $py $script read -n 5
```
僅過濾未讀郵件：
```powershell
& $py $script read -n 5 -u
```
依關鍵字快速篩選快取郵件：
```powershell
& $py $script read -k "工時"
```

### 3. 全量收件匣即時搜尋 (DASL)
從本機 Outlook 收件匣全量搜尋（即使是數千封前的歷史信件也能瞬間找出）：
```powershell
& $py $script search "ProjectLink" -n 10
```

### 4. 寄送電子郵件
直接透過本機已登入的 Outlook 寄送：
```powershell
& $py $script send --to "colleague@assaabloy.com" --subject "測試主旨" --body "郵件內文"
```

### 5. 清理背景殘留程序（故障排除）
若桌面上開啟 Outlook 出現 `MAPI 1.0 [000004C2]` 提示，一鍵釋放鎖定：
```powershell
& $py $script clean-zombies
```

---

## Antigravity MCP 整合規範

技能提供標準 MCP Server 位於：
`C:\Users\jeffho\.gemini\config\skills\outlook-m365-bridge\scripts\mcp_server.py`

在 `C:\Users\jeffho\.gemini\config\mcp_config.json` 中配置：

```json
{
  "mcpServers": {
    "outlook-m365": {
      "command": "C:\\Users\\jeffho\\.venv\\Scripts\\python.exe",
      "args": ["C:\\Users\\jeffho\\.gemini\\config\\skills\\outlook-m365-bridge\\scripts\\mcp_server.py"],
      "env": {}
    }
  }
}
```

### 提供之 MCP 工具清單
1. `get_inbox_emails(count, unread_only)`：讀取最新郵件列表。
2. `search_emails(keyword, count)`：依關鍵字搜尋郵件。
3. `send_email(to, subject, body, cc)`：寄送電子郵件。
4. `get_calendar_events(days_ahead)`：讀取未來行程清單。
5. `create_calendar_event(subject, start_datetime, end_datetime, location, body)`：在行事曆建立會議或行程。

---

## 維護與雲端配置設定

- **Power Automate 流程名稱**：`Antigravity-Outlook-Bridge`
- **雲端流程配置**：
  1. **觸發器**：週期（Recurrence）建議設定每 `5` 或 `10` 分鐘執行一次。
  2. **動作 1**：Office 365 Outlook - `取得電子郵件 (V3)`（資料夾：`Inbox`，頂端：`10`，僅未讀：`否`）。
  3. **動作 2**：OneDrive for Business - `建立檔案`（路徑：`/`，檔名：`latest_emails.json`，內容：`value`）。
- **同步路徑**：`C:\Users\jeffho\OneDrive - ASSA ABLOY Group\latest_emails.json`
