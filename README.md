# 📬 Outlook M365 Bridge for Antigravity & AI Agents

> 專為嚴格企業資安環境（如 ASSA ABLOY / 東隆五金 TLM）設計之 Outlook M365 雙軌郵件與行事曆整合技能。

[![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D6.svg)](https://microsoft.com)
[![Protocol](https://img.shields.io/badge/protocol-MCP%20(Model%20Context%20Protocol)-8A2BE2.svg)](https://modelcontextprotocol.io)

---

## 🌟 核心架構突破

在多數跨國企業或集團環境下，AI Agent 整合 Outlook 往往會遭遇多重資安封鎖：
1. **Azure AD 401 封鎖**：員工無權建立 App Registration，無法直接申請 Microsoft Graph API 權限。
2. **GPO 封鎖 COM 屬性**：透過 win32com 讀取郵件內文會遭到 Object Model Guard 阻擋（報錯 `0x80004004` 或屬性空白）。
3. **MAPI 鎖定卡死**：背景子程序呼叫 Outlook 會導致桌面版開啟時出現 `MAPI 1.0 [000004C2]` 逾時錯誤。

### 💡 雙軌架構解決方案 (Dual-Track Architecture)

```
┌──────────────────────────────────────────────────────────────┐
│                    雙軌整合架構 (Dual-Track)                  │
├───────────────────────────────┬──────────────────────────────┤
│ 🔵 雲端讀信軌 (Cloud Read)    │ 🟢 本機寫入軌 (Local Write)   │
│                               │                              │
│ Office 365 Outlook (雲端)     │ 本機 Outlook Desktop         │
│          │                    │          │                   │
│ Power Automate 標準排程 (免付費) │ PowerShell COM 自動化        │
│          │                    │          │                   │
│ OneDrive latest_emails.json   │ 寄送信件 / 新增行程 / 全量搜尋│
│          │                    │                              │
└──────────┼────────────────────┴──────────┬───────────────────┘
           ▼                               ▼
    ┌─────────────────────────────────────────────┐
    │       Antigravity AI Agent / MCP Server     │
    └─────────────────────────────────────────────┘
```

---

## 💻 另一台電腦快速安裝教學

### 第一步：複製技能專案庫

在另一台電腦的 PowerShell 中執行（直接複製至 Antigravity 技能目錄）：

```powershell
# 建立 skills 目錄（若不存在）
New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.gemini\config\skills"

# Clone 本專案
git clone https://github.com/JeffreyHo0520/outlook-m365-bridge.git "$env:USERPROFILE\.gemini\config\skills\outlook-m365-bridge"
```

### 第二步：安裝 Python 依賴

```powershell
pip install pywin32 mcp
```

### 第三步：設定 Antigravity MCP 配置

在 `$env:USERPROFILE\.gemini\config\mcp_config.json` 加入此伺服器：

```json
{
  "mcpServers": {
    "outlook-m365": {
      "command": "python",
      "args": [
        "C:\\Users\\<你的使用者名稱>\\.gemini\\config\\skills\\outlook-m365-bridge\\scripts\\mcp_server.py"
      ],
      "env": {}
    }
  }
}
```

### 第四步：設定 Power Automate 雲端同步流程（一次性）

1. 登入 [Power Automate 網頁版](https://make.powerautomate.com)（使用公司 M365 帳號）。
2. 建立「已排程的雲端流程」，命名為 `Antigravity-Outlook-Bridge`：
   - **觸發器**：週期（每 5 ~ 10 分鐘執行一次）。
   - **步驟一**：Office 365 Outlook - `取得電子郵件 (V3)`（資料夾：`Inbox`，頂端：`10`，僅未讀：`否`）。
   - **步驟二**：OneDrive for Business - `建立檔案`（路徑：`/`，檔名：`latest_emails.json`，內容：選步驟一的動態內容 `value` / `值`）。
3. 儲存並在選單中點擊 **「開啟」**。

---

## 🛠️ 命令列工具 (CLI) 指南

您也可以直接使用腳本執行獨立操作：

```powershell
$script = "$env:USERPROFILE\.gemini\config\skills\outlook-m365-bridge\scripts\outlook_bridge.py"

# 1. 檢查同步狀態與 Outlook 程序健康度
python $script status

# 2. 讀取最新收件匣郵件（預設 5 封）
python $script read -n 5

# 3. 僅篩選未讀郵件
python $script read -u

# 4. 全量收件匣即時搜尋關鍵字（直接透過 DASL 即時過濾歷史所有信件）
python $script search "ProjectLink"

# 5. 寄送電子郵件
python $script send --to "target@company.com" --subject "測試" --body "郵件內文"

# 6. 一鍵清理背景殘留程序（避免 MAPI 1.0 鎖定報錯）
python $script clean-zombies
```

---

## 🤖 支援之 MCP 原生工具清單

| 工具名稱 | 參數 | 說明 |
|---|---|---|
| `get_inbox_emails` | `count`, `unread_only` | 讀取收件匣最新郵件（含寄件人、主旨、時間、未讀標記、內文摘要與附件） |
| `search_emails` | `keyword`, `count` | 依關鍵字搜尋收件匣郵件 |
| `send_email` | `to`, `subject`, `body`, `cc` | 透過本機 Outlook 快速寄信 |
| `get_calendar_events` | `days_ahead` | 讀取未來行程清單 |
| `create_calendar_event` | `subject`, `start_datetime`, `end_datetime`, `location`, `body` | 建立行事曆行程 |

---

## 📄 License
MIT License.
