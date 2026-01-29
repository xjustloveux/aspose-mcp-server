# Aspose MCP Server

[![GitHub release](https://img.shields.io/github/v/release/xjustloveux/aspose-mcp-server?include_prereleases)](https://github.com/xjustloveux/aspose-mcp-server/releases)
[![GitHub license](https://img.shields.io/github/license/xjustloveux/aspose-mcp-server?cacheSeconds=3600)](LICENSE)
[![.NET Version](https://img.shields.io/badge/.NET-8.0-512BD4)](https://dotnet.microsoft.com/)
[![Build Status](https://img.shields.io/github/actions/workflow/status/xjustloveux/aspose-mcp-server/build-multi-platform.yml?branch=master&label=build)](https://github.com/xjustloveux/aspose-mcp-server/actions/workflows/build-multi-platform.yml)
[![Test Status](https://img.shields.io/github/actions/workflow/status/xjustloveux/aspose-mcp-server/test.yml?branch=master&label=tests)](https://github.com/xjustloveux/aspose-mcp-server/actions/workflows/test.yml)
[![Test Coverage](https://codecov.io/gh/xjustloveux/aspose-mcp-server/branch/master/graph/badge.svg)](https://codecov.io/gh/xjustloveux/aspose-mcp-server)
[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=xjustloveux_aspose-mcp-server&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=xjustloveux_aspose-mcp-server)
[![Maintainability Rating](https://sonarcloud.io/api/project_badges/measure?project=xjustloveux_aspose-mcp-server&metric=sqale_rating)](https://sonarcloud.io/summary/new_code?id=xjustloveux_aspose-mcp-server)
[![MCP Version](https://img.shields.io/badge/MCP-2025--11--25-blue)](https://modelcontextprotocol.io/)
[![MCP SDK](https://img.shields.io/badge/MCP%20SDK-0.6.0-purple)](https://github.com/modelcontextprotocol/csharp-sdk)
[![Aspose Version](https://img.shields.io/badge/Aspose-23.10.0-orange)](https://products.aspose.com/total/net/)
[![xUnit](https://img.shields.io/badge/xUnit-2.9.2-blue?logo=xunit)](https://xunit.net/)

**格式支援：** ![Word](https://img.shields.io/badge/Word-DOC%2CDOCX-blue) ![Excel](https://img.shields.io/badge/Excel-XLS%2CXLSX-green) ![PowerPoint](https://img.shields.io/badge/PowerPoint-PPT%2CPPTX-orange) ![PDF](https://img.shields.io/badge/PDF-PDF-red)

基於 .NET 8.0 和 Aspose.Total 的 Model Context Protocol (MCP) 伺服器，為 MCP 客戶端提供強大的辦公文檔處理能力。

## ✨ 特性

### 核心功能
- **88 個統一工具** - Word(24)、Excel(25)、PowerPoint(21)、PDF(15)、轉換(2)、Session(1) 已整合
- **按需啟用** - 只啟用需要的文檔類型，減少資源佔用
- **跨平台** - Windows、Linux、macOS (Intel + ARM)，單一可執行檔案
- **開箱即用** - 預編譯版本無需安裝 .NET Runtime
- **完整讀寫** - 支援從A文檔讀取格式應用到B文檔

### 傳輸模式
- **Stdio 模式** (預設) - 標準輸入輸出，適用於本地 MCP 客戶端
- **HTTP 模式** - Streamable HTTP（MCP 2025-03-26+），適用於網頁應用
- **WebSocket 模式** - 雙向通訊，適用於即時互動

### 進階功能
- **Session 管理** - 在記憶體中編輯文件，支援 open/save/close 操作，支援多租戶隔離
- **認證機制** - 可選的 API Key 和 JWT 認證（4 種驗證模式）
- **追蹤系統** - 結構化日誌、Webhook 通知、Prometheus Metrics
- **Origin 驗證** - 防止 DNS 重綁定攻擊（HTTP/WebSocket 模式）

### 技術特性
- **MCP SDK 0.6.0** - 使用官方 ModelContextProtocol NuGet 套件，支援 Tool Annotations 和 outputSchema
- **Tool Annotations** - 所有工具標註 ReadOnly、Destructive、Idempotent、OpenWorld 行為特性
- **結構化輸出** - Handler 返回強型別結果，SDK 自動生成 outputSchema（oneOf JSON Schema）
- **統一字型設定** - 多個工具支援中英文字型分別設定（`fontNameAscii` 和 `fontNameFarEast` 參數）
- **靈活的授權配置** - 支援總授權或單一組件授權，自動搜尋、環境變數或命令列參數配置
- **安全加固** - 全面的路徑驗證、輸入驗證和錯誤處理

## 📑 目錄

**開始使用**
- [🚀 快速開始](#-快速開始) - 下載、配置、啟動
- [📦 功能概覽](#-功能概覽) - Word、Excel、PowerPoint、PDF、轉換工具
- [📋 工具列表](#-工具列表) - 88 個工具的詳細說明

**進階配置**
- [🔌 傳輸模式](#-傳輸模式) - Stdio、HTTP、WebSocket 配置
- [📂 Session 管理](#-session-管理) - 文件記憶體編輯、DocumentSessionTool
- [🔐 認證機制](#-認證機制) - API Key、JWT 認證配置
- [📡 追蹤系統](#-追蹤系統) - 日誌、Webhook、Prometheus Metrics
- [🚢 部署指南](#-部署指南) - Docker、Kubernetes、IIS 部署

**開發與技術**
- [🛠️ 開發者指南](#️-開發者指南) - 倉庫結構、本地開發、多平台構建、運行測試
- [🔒 安全特性](#-安全特性) - 路徑驗證、輸入驗證、錯誤處理
- [🌍 跨平台支援](#-跨平台支援) - Windows、Linux、macOS 技術規格、Linux/macOS 額外需求

**參考資料**
- [📝 使用範例](#-使用範例) - 從A文檔複製格式到B文檔
- [⚠️ 重要說明](#️-重要說明) - 索引行為、參數命名一致性
- [📄 授權](#-授權) - Aspose 授權配置方式
- [❓ 常見問題](#-常見問題) - FAQ

**其他**
- [🔗 相關資源](#-相關資源) - 官方文檔、MCP 客戶端、專案資源
- [📊 專案統計](#-專案統計) - 工具數、測試覆蓋率、技術規格

## 🚀 快速開始

### 1. 下載預編譯版本

從 [GitHub Releases](https://github.com/xjustloveux/aspose-mcp-server/releases) 下載最新版本：
- Windows: `aspose-mcp-server-windows-x64.zip`
- Linux: `aspose-mcp-server-linux-x64.zip`
- macOS Intel: `aspose-mcp-server-macos-x64.zip`
- macOS ARM: `aspose-mcp-server-macos-arm64.zip`

解壓到任意目錄，例如：
- Windows: `C:\Tools\aspose-mcp-server\`
- macOS/Linux: `~/tools/aspose-mcp-server/`

**放置授權檔案：** 將授權檔案放在可執行檔案同一目錄。支援以下方式：

- **總授權**：`Aspose.Total.lic`（包含所有組件）
- **單一組件授權**：`Aspose.Words.lic`、`Aspose.Cells.lic`、`Aspose.Slides.lic`、`Aspose.Pdf.lic`
- **自訂檔案名稱**：可透過環境變數或命令列參數指定

**授權檔案配置方式：**

1. **自動搜尋**（推薦）：將授權檔案放在可執行檔案目錄，系統會自動搜尋
2. **環境變數**：設定 `ASPOSE_LICENSE_PATH` 環境變數指向授權檔案路徑
3. **命令列參數**：使用 `--license 路徑` 指定授權檔案

**範例：**
```json
{
  "mcpServers": {
    "aspose-word": {
      "command": "C:/Tools/aspose-mcp-server/AsposeMcpServer.exe",
      "args": ["--word", "--license", "C:/Licenses/Aspose.Words.lic"]
    }
  }
}
```

**注意**：如果找不到授權檔案，系統會以試用模式運行（會有試用版標記）。

### 2. 配置 MCP 客戶端

根據您使用的 MCP 客戶端，編輯對應的配置檔案。配置檔案通常位於應用程式的設定目錄中，請參考您使用的客戶端文檔以確認具體路徑。

**配置範例：**
```json
{
  "mcpServers": {
    "aspose-word": {
      "command": "C:/Tools/aspose-mcp-server/AsposeMcpServer.exe",
      "args": ["--word"]
    }
  }
}
```

**可用參數**:
- `--word` - Word 工具（自動包含轉換功能）
- `--excel` - Excel 工具（自動包含轉換功能）
- `--powerpoint` / `--ppt` - PowerPoint 工具（自動包含轉換功能）
- `--pdf` - PDF 工具
- `--all` - 所有工具
- `--session-enabled` - 啟用 Session 管理（`document_session` 工具）
- `--license 路徑` - 指定授權檔案路徑（可選）

> **工具過濾**：只有啟用的工具類別會出現在 MCP 工具列表中。例如使用 `--word` 時，只會顯示 `word_*` 相關工具。

**轉換功能說明**：
- 啟用任何文檔工具（`--word`、`--excel`、`--ppt`）時，自動包含 `convert_to_pdf`（轉換為PDF）
- 啟用兩個或以上文檔工具時，自動包含 `convert_document`（跨格式轉換，如Word轉Excel）

📋 **更多配置範例：** `config_example.json`（配置格式適用於所有 MCP 客戶端）

### 3. 重啟 MCP 客戶端

完成配置後，重啟您使用的 MCP 客戶端（如 Claude Desktop、Cursor 等）即可開始使用。

## 📦 功能概覽

### Word (24個工具)

**檔案操作**：創建、讀取、轉換、合併、拆分、從範本創建  
**內容編輯**：文字（中英文字型分別設定）、段落、表格、圖片、圖表、清單、文字框、欄位  
**格式設定**：段落格式、字型、樣式（複製樣式保留字型）、頁首頁尾、頁面設定  
**高級功能**：書籤、超連結、註釋、目錄、文檔屬性、保護、郵件合併、腳註、尾註、交叉引用、索引、文檔比較、修訂管理、表單欄位、水印、形狀

### Excel (25個工具)

**檔案操作**：創建、讀取、寫入、轉換、合併、拆分、保護  
**工作表操作**：添加、刪除、重新命名、複製、移動、隱藏/顯示、讀取資訊  
**行列操作**：插入/刪除行/列、設定行高列寬  
**單元格操作**：合併/取消合併、插入/刪除單元格、鎖定/解鎖  
**資料操作**：排序、篩選、資料驗證、查找替換、批次寫入  
**格式設定**：單元格格式、條件格式、樣式  
**高級功能**：圖表、公式、資料透視表、凍結窗格、超連結、圖片、頁面設定、陣列公式、列印設定、工作表外觀設定、分組、命名範圍

### PowerPoint (21個工具)

**檔案操作**：創建、讀取、轉換、合併、拆分
**投影片管理**：添加、刪除、移動、複製、隱藏、設定版面配置、設定大小
**內容編輯**：文字、圖片、表格、圖表、形狀、SmartArt、媒體（音訊/影片）
**格式設定**：文字格式、形狀格式、背景、頁眉頁腳、主題
**高級功能**：動畫、轉場、備註、章節、超連結、文檔屬性

### PDF (15個工具)

**檔案操作**：創建、讀取、合併、拆分、壓縮、加密  
**內容添加**：文字、圖片、表格、水印、頁面、書籤、註釋、連結、表單欄位、附件  
**編輯操作**：編輯文字、表格、書籤、註釋、連結、表單欄位、圖片  
**讀取操作**：提取文字、圖片、讀取頁面資訊、書籤、註釋、連結、表單欄位、附件、簽名、統計資訊  
**高級功能**：簽名、頁面旋轉、編輯（塗黑）

### 轉換工具 (2個)

- `convert_to_pdf` - 將任何文檔轉換為PDF（啟用任何文檔工具時自動可用）
- `convert_document` - 跨格式轉換（啟用兩個或以上文檔工具時自動可用）

## 🔌 傳輸模式

Aspose MCP Server 支援三種傳輸模式，可透過命令列參數或環境變數配置：

### Stdio 模式（預設）

標準輸入輸出模式，適用於本地 MCP 客戶端整合：

```bash
# 預設使用 Stdio 模式
AsposeMcpServer.exe --word

# 或明確指定 Stdio 模式
AsposeMcpServer.exe --stdio --word
```

### HTTP 模式

Streamable HTTP 模式（MCP 2025-03-26+），適用於網頁應用：

```bash
# 命令列參數
AsposeMcpServer.exe --http --port 3000 --word

# 或使用環境變數
set ASPOSE_TRANSPORT=http
set ASPOSE_PORT=3000
AsposeMcpServer.exe --word
```

連接端點：`http://localhost:3000/mcp`

### WebSocket 模式

雙向通訊模式，適用於即時互動應用：

```bash
# 命令列參數
AsposeMcpServer.exe --ws --port 3000 --word

# 或使用 --websocket 別名
AsposeMcpServer.exe --websocket --port 3000 --word

# 或使用環境變數
set ASPOSE_TRANSPORT=ws
set ASPOSE_PORT=3000
AsposeMcpServer.exe --word
```

連接端點：`ws://localhost:3000/ws`

### 傳輸模式環境變數

| 變數 | 說明 | 預設值 |
|------|------|--------|
| `ASPOSE_TRANSPORT` | 傳輸模式 (stdio/http/ws) | stdio |
| `ASPOSE_PORT` | 監聽埠號（1-65535，無效值重設為 3000） | 3000 |
| `ASPOSE_HOST` | 監聽位址（`localhost`、`0.0.0.0`、`*` 或特定 IP，無效值重設為 localhost） | localhost |
| `ASPOSE_TOOLS` | 啟用的工具 (all 或 word,excel,pdf,ppt) | all |

> **注意**:
> - Docker/Kubernetes 部署時需設定 `ASPOSE_HOST=0.0.0.0` 以便容器外部可以訪問
> - `localhost` 僅允許本機連接，`0.0.0.0` 或 `*` 允許所有介面連接

## 📂 Session 管理

Session 管理功能允許在記憶體中編輯文件，避免頻繁的磁碟讀寫操作。

> **注意**：Session 功能預設停用。必須使用 `--session-enabled` 參數或設定 `ASPOSE_SESSION_ENABLED=true` 環境變數才會啟用 `document_session` 工具。

### DocumentSessionTool

使用 `document_session` 工具管理文件 Session：

```
# 開啟文件到記憶體
document_session(operation="open", path="document.docx", sessionId="my-session")

# 在記憶體中編輯（使用其他工具時指定 sessionId）
word_text(operation="add", sessionId="my-session", text="新增內容")

# 儲存變更到磁碟
document_session(operation="save", sessionId="my-session")

# 另存為新檔案
document_session(operation="save", sessionId="my-session", outputPath="new-document.docx")

# 關閉 Session（不儲存）
document_session(operation="close", sessionId="my-session")

# 列出所有 Session
document_session(operation="list")
```

### Session 操作

| 操作 | 說明 |
|------|------|
| `open` | 開啟文件到記憶體 Session |
| `save` | 儲存 Session 變更到原始檔案 |
| `close` | 關閉 Session（可選擇儲存或捨棄變更） |
| `list` | 列出所有活動 Session |
| `status` | 取得特定 Session 狀態 |

### 暫存檔操作

當 Session 結束時，根據 `ASPOSE_SESSION_ON_DISCONNECT` 設定處理未儲存的變更（預設 `SaveToTemp`）。

**觸發時機：**

| 情境 | 說明 |
|------|------|
| 閒置逾時 | 依據 `ASPOSE_SESSION_TIMEOUT` 設定（預設 30 分鐘），最常見的觸發情境 |
| 手動關閉 | 呼叫 `document_session(operation="close")` 時 |
| 伺服器優雅關閉 | 正常終止程序時（SIGTERM 等） |

> **注意**：強制終止（如 `kill -9`）不會觸發此設定，建議啟用 `ASPOSE_SESSION_AUTO_SAVE_INTERVAL` 定期暫存來保護資料。

使用以下操作管理暫存檔：

| 操作 | 說明 |
|------|------|
| `list_temp` | 列出可恢復的暫存檔 |
| `recover` | 恢復暫存檔到原始路徑或指定路徑 |
| `delete_temp` | 刪除特定暫存檔 |
| `cleanup` | 清理過期暫存檔 |
| `temp_stats` | 取得暫存檔統計資訊 |

**暫存檔恢復範例：**
```
# 列出可恢復的暫存檔
document_session(operation="list_temp")

# 恢復到原始路徑
document_session(operation="recover", sessionId="sess_abc123")

# 恢復到指定路徑
document_session(operation="recover", sessionId="sess_abc123", outputPath="recovered.docx")
```

### Session 配置

**環境變數：**

| 變數 | 說明 | 預設值 |
|------|------|--------|
| `ASPOSE_SESSION_ENABLED` | 啟用 Session 管理 | false |
| `ASPOSE_SESSION_MAX` | 最大 Session 數 | 10 |
| `ASPOSE_SESSION_TIMEOUT` | 閒置超時（分鐘） | 30 |
| `ASPOSE_SESSION_MAX_FILE_SIZE_MB` | 最大檔案大小（MB，最小 1） | 100 |
| `ASPOSE_SESSION_TEMP_DIR` | 臨時目錄 | 系統臨時目錄 |
| `ASPOSE_SESSION_TEMP_RETENTION_HOURS` | 暫存檔保留時間（小時） | 24 |
| `ASPOSE_SESSION_ON_DISCONNECT` | 逾時處理 (AutoSave/Discard/SaveToTemp/PromptOnReconnect) | SaveToTemp |
| `ASPOSE_SESSION_ISOLATION` | 隔離模式 (none/group) | group |
| `ASPOSE_SESSION_AUTO_SAVE_INTERVAL` | 自動暫存間隔（分鐘，0 = 停用） | 0 |

**自動暫存說明：**
- 當設定為大於 0 的值時，系統會定期將有變更的 Session 暫存到臨時目錄
- 此功能可防止因非預期終止（如 kill -9）導致的資料遺失
- 暫存檔會保存到 `ASPOSE_SESSION_TEMP_DIR` 指定的目錄
- 自動暫存**不會**關閉或影響活動中的 Session，僅作為資料保護機制

**隔離模式說明：**
- `none` - 無隔離，所有用戶可存取所有 session（Stdio 模式向後兼容）
- `group` - 群組級隔離，同群組內的用戶可相互存取 session（預設）。群組識別符來源可透過 `ASPOSE_AUTH_JWT_GROUP_CLAIM` 配置（預設為 tenant_id）

**命令行參數：**

| 參數 | 說明 |
|------|------|
| `--session-enabled` | 啟用 Session 管理 |
| `--session-disabled` | 停用 Session 管理 |
| `--session-max N` | 最大 Session 數 |
| `--session-timeout N` | 閒置超時（分鐘） |
| `--session-max-file-size N` | 最大檔案大小（MB，最小 1） |
| `--session-temp-dir path` | 臨時目錄 |
| `--session-temp-retention-hours N` | 暫存檔保留時間（小時） |
| `--session-on-disconnect behavior` | 逾時處理 |
| `--session-isolation mode` | 隔離模式 (none/group) |
| `--session-auto-save N` | 自動暫存間隔（分鐘，0 = 停用） |

## 🔐 認證機制

啟用 HTTP 或 WebSocket 模式時，可配置認證機制保護 API：

> **雙重認證模式**：當 API Key 和 JWT 同時啟用時，請求必須**同時通過兩者驗證**（串聯模式）。API Key 驗證先執行，失敗則直接返回 401；通過後再執行 JWT 驗證。若只需其中一種認證通過，請僅啟用其一。

### API Key 認證

支援四種驗證模式：

**Local 模式**（本地驗證）：
```bash
set ASPOSE_AUTH_APIKEY_ENABLED=true
set ASPOSE_AUTH_APIKEY_MODE=local
set ASPOSE_AUTH_APIKEY_KEYS=key1:group1,key2:group2
```

**Gateway 模式**（信任 API Gateway 傳遞的標頭）：
```bash
set ASPOSE_AUTH_APIKEY_ENABLED=true
set ASPOSE_AUTH_APIKEY_MODE=gateway
# 從 X-API-Key 和 X-Group-Id 標頭讀取
```

**Introspection 模式**（呼叫外部端點驗證）：
```bash
set ASPOSE_AUTH_APIKEY_ENABLED=true
set ASPOSE_AUTH_APIKEY_MODE=introspection
set ASPOSE_AUTH_APIKEY_INTROSPECTION_URL=https://auth.example.com/validate
```

端點需回傳 JSON：`{"active": true, "group_id": "..."}`

**Custom 模式**（自訂驗證邏輯）：
```bash
set ASPOSE_AUTH_APIKEY_ENABLED=true
set ASPOSE_AUTH_APIKEY_MODE=custom
set ASPOSE_AUTH_APIKEY_CUSTOM_URL=https://auth.example.com/custom
```

端點需回傳 JSON：`{"valid": true, "group_id": "..."}`

### JWT 認證

支援四種驗證模式：

**Local 模式**（本地驗證 JWT）：
```bash
set ASPOSE_AUTH_JWT_ENABLED=true
set ASPOSE_AUTH_JWT_MODE=local
set ASPOSE_AUTH_JWT_SECRET=your-256-bit-secret
set ASPOSE_AUTH_JWT_ISSUER=your-issuer
set ASPOSE_AUTH_JWT_AUDIENCE=your-audience
```

**使用 RSA 公鑰驗證**：
```bash
set ASPOSE_AUTH_JWT_ENABLED=true
set ASPOSE_AUTH_JWT_MODE=local
set ASPOSE_AUTH_JWT_PUBLIC_KEY_PATH=/path/to/public.pem
```

**Gateway/Introspection/Custom 模式**：與 API Key 類似配置

JWT Introspection 端點需回傳 JSON：`{"active": true, "group_id": "...", "sub": "..."}`

JWT Custom 端點需回傳 JSON：`{"valid": true, "group_id": "...", "user_id": "..."}`

### 外部端點請求格式

MCP Server 發送到外部驗證端點的請求格式：

| 模式 | Content-Type | 請求格式 |
|------|-------------|---------|
| API Key Introspection | `application/x-www-form-urlencoded` | `{IntrospectionKeyField}=<apiKey>` |
| API Key Custom | `application/json` | `{"apiKey": "..."}` |
| JWT Introspection (RFC 7662) | `application/x-www-form-urlencoded` | `token=<token>&token_type_hint=access_token` |
| JWT Custom | `application/json` | `{"token": "..."}`|

### 同時啟用多種認證

若同時啟用 API Key 和 JWT 認證，請求必須**依序通過兩者驗證**才算成功：

```
請求 → API Key 驗證 → JWT 驗證 → 處理請求
         ↓ 失敗          ↓ 失敗
        401             401
```

**適用情境**：
- API Key 識別應用程式/服務，JWT 識別使用者
- 多租戶環境中的雙重身份驗證

**注意**：大多數情況只需啟用其中一種認證方式即可。

### 認證配置

**環境變數：**

| 變數 | 說明 | 預設值 |
|------|------|--------|
| `ASPOSE_AUTH_APIKEY_ENABLED` | 啟用 API Key 認證 | false |
| `ASPOSE_AUTH_APIKEY_MODE` | 模式 (local/gateway/introspection/custom) | local |
| `ASPOSE_AUTH_APIKEY_KEYS` | API Key 列表 (key:group,...) | - |
| `ASPOSE_AUTH_APIKEY_HEADER` | API Key 標頭名稱 | X-API-Key |
| `ASPOSE_AUTH_APIKEY_GROUP_HEADER` | 群組 ID 標頭名稱 (Gateway 模式) | X-Group-Id |
| `ASPOSE_AUTH_APIKEY_INTROSPECTION_AUTH` | Introspection 認證標頭值 | - |
| `ASPOSE_AUTH_APIKEY_TIMEOUT` | 外部驗證逾時（秒，Introspection/Custom） | 5 |
| `ASPOSE_AUTH_APIKEY_CACHE_ENABLED` | 啟用驗證結果快取 | true |
| `ASPOSE_AUTH_APIKEY_CACHE_TTL` | 快取存活時間（秒） | 300 |
| `ASPOSE_AUTH_APIKEY_CACHE_MAX_SIZE` | 快取最大項目數 | 10000 |
| `ASPOSE_AUTH_JWT_ENABLED` | 啟用 JWT 認證 | false |
| `ASPOSE_AUTH_JWT_MODE` | 模式 (local/gateway/introspection/custom) | local |
| `ASPOSE_AUTH_JWT_SECRET` | HMAC 密鑰 | - |
| `ASPOSE_AUTH_JWT_ISSUER` | 預期發行者 | - |
| `ASPOSE_AUTH_JWT_AUDIENCE` | 預期受眾 | - |
| `ASPOSE_AUTH_JWT_GROUP_CLAIM` | 群組 ID Claim 名稱 (如 tenant_id, team_id, org_id) | tenant_id |
| `ASPOSE_AUTH_JWT_USER_CLAIM` | 使用者 ID Claim 名稱 | sub |
| `ASPOSE_AUTH_JWT_GROUP_HEADER` | 群組 ID 標頭名稱 (Gateway 模式) | X-Group-Id |
| `ASPOSE_AUTH_JWT_USER_HEADER` | 使用者 ID 標頭名稱 (Gateway 模式) | X-User-Id |
| `ASPOSE_AUTH_JWT_CLIENT_ID` | OIDC 客戶端 ID (Introspection 模式) | - |
| `ASPOSE_AUTH_JWT_CLIENT_SECRET` | OIDC 客戶端密鑰 (Introspection 模式) | - |
| `ASPOSE_AUTH_JWT_TIMEOUT` | 外部驗證逾時（秒，Introspection/Custom） | 5 |
| `ASPOSE_AUTH_JWT_CACHE_ENABLED` | 啟用驗證結果快取 | true |
| `ASPOSE_AUTH_JWT_CACHE_TTL` | 快取存活時間（秒） | 300 |
| `ASPOSE_AUTH_JWT_CACHE_MAX_SIZE` | 快取最大項目數 | 10000 |

**命令行參數：**

| 參數 | 說明 |
|------|------|
| `--auth-apikey-enabled` | 啟用 API Key 認證 |
| `--auth-apikey-disabled` | 停用 API Key 認證 |
| `--auth-apikey-mode mode` | API Key 驗證模式 |
| `--auth-apikey-keys key1:group1,key2:group2` | API Key 列表（group 可包含冒號） |
| `--auth-apikey-header name` | API Key 標頭名稱 |
| `--auth-apikey-group-header name` | 群組 ID 標頭名稱 |
| `--auth-apikey-introspection-url url` | Introspection 端點 URL |
| `--auth-apikey-introspection-auth value` | Introspection 認證標頭值 |
| `--auth-apikey-introspection-field name` | Introspection 請求欄位名稱（預設：key） |
| `--auth-apikey-custom-url url` | Custom 驗證端點 URL |
| `--auth-apikey-timeout N` | 外部驗證逾時（秒，Introspection/Custom） |
| `--auth-apikey-cache-enabled` | 啟用驗證結果快取 |
| `--auth-apikey-cache-ttl N` | 快取存活時間（秒） |
| `--auth-apikey-cache-max-size N` | 快取最大項目數 |
| `--auth-jwt-enabled` | 啟用 JWT 認證 |
| `--auth-jwt-disabled` | 停用 JWT 認證 |
| `--auth-jwt-mode mode` | JWT 驗證模式 |
| `--auth-jwt-secret value` | HMAC 密鑰 |
| `--auth-jwt-public-key-path path` | RSA/ECDSA 公鑰文件路徑 |
| `--auth-jwt-issuer value` | 預期發行者 |
| `--auth-jwt-audience value` | 預期受眾 |
| `--auth-jwt-group-claim name` | 群組 ID Claim 名稱 (如 tenant_id, team_id, org_id) |
| `--auth-jwt-user-claim name` | 使用者 ID Claim 名稱 |
| `--auth-jwt-group-header name` | 群組 ID 標頭名稱 |
| `--auth-jwt-user-header name` | 使用者 ID 標頭名稱 |
| `--auth-jwt-introspection-url url` | OAuth Introspection 端點 URL |
| `--auth-jwt-client-id value` | OAuth Client ID（Introspection 模式） |
| `--auth-jwt-client-secret value` | OAuth Client Secret（Introspection 模式） |
| `--auth-jwt-custom-url url` | Custom 驗證端點 URL |
| `--auth-jwt-timeout N` | 外部驗證逾時（秒，Introspection/Custom） |
| `--auth-jwt-cache-enabled` | 啟用驗證結果快取 |
| `--auth-jwt-cache-ttl N` | 快取存活時間（秒） |
| `--auth-jwt-cache-max-size N` | 快取最大項目數 |

## 📡 追蹤系統

追蹤系統提供多種監控和日誌記錄方式：

### 結構化日誌

```bash
set ASPOSE_LOG_ENABLED=true
set ASPOSE_LOG_TARGETS=Console,EventLog
```

支援的日誌目標：
- `Console` - 輸出到 stderr（遵循 MCP 規範）
- `EventLog` - Windows 事件日誌（僅限 Windows）

> **注意**：如需將日誌寫入檔案，建議使用外部工具（如 Docker 日誌驅動、systemd journal、IIS stdout 日誌）進行日誌收集。

### Webhook 通知

每次工具呼叫時發送 HTTP POST 通知：

```bash
set ASPOSE_WEBHOOK_ENABLED=true
set ASPOSE_WEBHOOK_URL=https://your-server.com/webhook
```

Webhook 載荷格式：
```json
{
  "timestamp": "2025-01-01T12:00:00Z",
  "groupId": "group1",
  "tool": "word_text",
  "operation": "add",
  "durationMs": 150,
  "success": true,
  "errorMessage": null
}
```

### Prometheus Metrics

暴露 `/metrics` 端點供 Prometheus 抓取：

```bash
set ASPOSE_METRICS_ENABLED=true
set ASPOSE_METRICS_PATH=/metrics
```

可用指標：
- `aspose_mcp_requests_total` - 總請求數（按工具、操作、狀態分類）
- `aspose_mcp_request_duration_seconds` - 請求處理時間

### 追蹤配置

**環境變數：**

| 變數 | 說明 | 預設值 |
|------|------|--------|
| `ASPOSE_LOG_ENABLED` | 啟用日誌 | true |
| `ASPOSE_LOG_TARGETS` | 日誌目標 (Console,EventLog) | Console |
| `ASPOSE_WEBHOOK_ENABLED` | 啟用 Webhook | false |
| `ASPOSE_WEBHOOK_URL` | Webhook URL | - |
| `ASPOSE_WEBHOOK_AUTH_HEADER` | Webhook 認證標頭值 | - |
| `ASPOSE_WEBHOOK_TIMEOUT` | Webhook 超時（1-300 秒，無效值重設為 5） | 5 |
| `ASPOSE_METRICS_ENABLED` | 啟用 Metrics | false |
| `ASPOSE_METRICS_PATH` | Metrics 路徑 | /metrics |

**命令行參數：**

| 參數 | 說明 |
|------|------|
| `--log-enabled` | 啟用日誌 |
| `--log-disabled` | 停用日誌 |
| `--log-targets Console,EventLog` | 日誌目標 |
| `--webhook-enabled` | 啟用 Webhook |
| `--webhook-disabled` | 停用 Webhook |
| `--webhook-url url` | Webhook URL（設定後自動啟用） |
| `--webhook-auth-header header` | Webhook 認證標頭 |
| `--webhook-timeout N` | Webhook 超時（1-300 秒，無效值重設為 5） |
| `--metrics-enabled` | 啟用 Metrics |
| `--metrics-disabled` | 停用 Metrics |
| `--metrics-path path` | Metrics 路徑 |

## 🚢 部署指南

Aspose MCP Server 支援多種部署方式：

### Docker 部署

```bash
# 建置映像
docker build -f deploy/Dockerfile -t aspose-mcp-server .

# 執行容器
docker run -d -p 3000:3000 \
  -e ASPOSE_TRANSPORT=http \
  -e ASPOSE_HOST=0.0.0.0 \
  -e ASPOSE_TOOLS=all \
  aspose-mcp-server
```

使用 Docker Compose：
```bash
docker-compose -f deploy/docker-compose.yml up -d
```

### Kubernetes 部署

```bash
# 建立 License Secret（可選）
kubectl create secret generic aspose-license \
  --from-file=Aspose.Total.lic=/path/to/license

# 部署應用程式
kubectl apply -f deploy/deployment.yaml
```

提供的範本包含 Deployment 配置，可根據需求自行添加 Service、Ingress 等資源。

### IIS 部署

```powershell
# 發布應用程式
dotnet publish -c Release -o ./publish

# 複製 web.config 到發布目錄
copy deploy/web.config ./publish/

# 在 IIS 中建立網站，指向 publish 目錄
```

前置需求：
- Windows Server 2019/2022
- IIS 10.0+
- .NET 8.0 Hosting Bundle
- WebSocket 功能已啟用（如使用 WebSocket 傳輸模式）

### Health Check 端點

HTTP/WebSocket 模式下提供以下端點：
- `GET /health` - 健康檢查
- `GET /ready` - 就緒檢查
- `GET /metrics` - Prometheus 指標（需啟用）

詳細部署說明請參閱 [部署指南](https://xjustloveux.github.io/aspose-mcp-server/deployment.html)。

## 🔒 安全特性

### Origin 驗證

HTTP 和 WebSocket 模式預設啟用 Origin 標頭驗證，防止 DNS 重綁定攻擊：

```bash
# 停用 Origin 驗證（不建議用於生產環境）
AsposeMcpServer.exe --http --no-origin-validation

# 不允許 localhost（生產環境）
AsposeMcpServer.exe --http --no-localhost

# 要求必須有 Origin 標頭
AsposeMcpServer.exe --http --require-origin

# 指定允許的 Origin 清單
AsposeMcpServer.exe --http --allowed-origins https://app.example.com,https://admin.example.com
```

**環境變數：**

| 變數 | 說明 | 預設值 |
|------|------|--------|
| `ASPOSE_ORIGIN_VALIDATION` | 啟用 Origin 驗證 | true |
| `ASPOSE_ALLOW_LOCALHOST` | 允許 localhost Origin | true |
| `ASPOSE_ALLOW_MISSING_ORIGIN` | 允許缺少 Origin 標頭 | true |
| `ASPOSE_ALLOWED_ORIGINS` | 允許的 Origin 清單（逗號分隔） | - |

**命令行參數：**

| 參數 | 說明 |
|------|------|
| `--no-origin-validation` | 停用 Origin 驗證 |
| `--no-localhost` | 不允許 localhost Origin |
| `--require-origin` | 要求必須有 Origin 標頭 |
| `--allowed-origins origins` | 允許的 Origin 清單（逗號分隔） |

### 路徑驗證
- ✅ 所有檔案路徑都經過 `SecurityHelper.ValidateFilePath()` 驗證
- ✅ 防止路徑遍歷攻擊（`../`, `..\`）
- ✅ 限制路徑長度（最大260字元）和檔案名稱長度（最大255字元）
- ✅ 驗證路徑中的非法字元

### 輸入驗證
- ✅ 陣列大小驗證（`SecurityHelper.ValidateArraySize`，最大1000項）
- ✅ 字串長度驗證（`SecurityHelper.ValidateStringLength`，最大10000字元）

### 錯誤處理
- ✅ 錯誤訊息清理，防止資訊洩露
- ✅ 移除檔案路徑、堆疊追蹤等敏感資訊
- ✅ 生產環境不暴露詳細錯誤資訊

### 安全限制
- **最大路徑長度**: 260 字元
- **最大檔案名稱長度**: 255 字元
- **最大陣列大小**: 1000 項
- **最大字串長度**: 10000 字元
- **預設不允許絕對路徑**: 否（可透過參數允許）

## 🛠️ 開發者指南

### 倉庫結構
```
aspose-mcp-server/
├── Tools/                    📁 工具原始碼（MCP Tool 入口點）
│   ├── Word/                 24 個工具
│   ├── Excel/                25 個工具
│   ├── PowerPoint/           21 個工具
│   ├── PDF/                  15 個工具
│   ├── Conversion/           2 個工具
│   └── Session/              1 個工具 (DocumentSessionTool)
├── Handlers/                 📁 操作處理器（業務邏輯實作）
│   ├── Word/                 Word 處理器
│   ├── Excel/                Excel 處理器
│   ├── PowerPoint/           PowerPoint 處理器
│   └── Pdf/                  PDF 處理器
├── Helpers/                  🛠️ 通用輔助工具
│   ├── Excel/                Excel 專用 Helper
│   ├── Word/                 Word 專用 Helper
│   ├── PowerPoint/           PowerPoint 專用 Helper
│   └── Pdf/                  PDF 專用 Helper
├── Results/                  📊 結果類型定義
│   ├── Common/               通用結果類型（SuccessResult 等）
│   ├── Word/                 Word 操作結果類型
│   ├── Excel/                Excel 操作結果類型
│   ├── PowerPoint/           PowerPoint 操作結果類型
│   └── Pdf/                  PDF 操作結果類型
├── Core/                     🔧 MCP 伺服器核心
│   ├── Handlers/             Handler 基礎架構
│   ├── Progress/             MCP Progress 支援（文檔處理進度回報）
│   ├── Security/             安全模組（認證、Origin 驗證）
│   ├── Session/              Session 管理模組
│   ├── Tracking/             追蹤模組（日誌、Webhook、Metrics）
│   ├── Transport/            傳輸層模組
│   └── ShapeDetailProviders/ PowerPoint 形狀詳細資訊提供者
├── Tests/                    🧪 單元測試與整合測試
│   ├── Core/                 核心功能測試
│   ├── Handlers/             Handler 測試
│   ├── Tools/                工具測試
│   ├── Helpers/              Helper 單元測試
│   ├── Integration/          整合測試
│   └── Infrastructure/       測試基礎設施（TestBase 等）
├── deploy/                   🚢 部署與構建配置
├── docs/                     📚 GitHub Pages 文檔
├── .github/workflows/        🔄 GitHub Actions 工作流程
└── bin/                      ❌ 本地編譯輸出（不在版本控制）
```

### 本地開發

```bash
# 複製倉庫
git clone https://github.com/xjustloveux/aspose-mcp-server.git
cd aspose-mcp-server

# 編譯 Release 版本
pwsh deploy/build.ps1 --configuration Release

# 發布 Windows 版本
pwsh deploy/publish.ps1 -Windows

# 發布所有平台
pwsh deploy/publish.ps1 -All
```

### 多平台構建

**本地構建：**
```bash
# Windows
pwsh deploy/publish.ps1 -Windows

# Linux
pwsh deploy/publish.ps1 -Linux

# macOS (Intel + ARM)
pwsh deploy/publish.ps1 -MacOS

# 所有平台
pwsh deploy/publish.ps1 -All

# 清理後構建
pwsh deploy/publish.ps1 -All -Clean
```

**構建產物位置：**
- Windows: `publish/windows-x64/AsposeMcpServer.exe`
- Linux: `publish/linux-x64/AsposeMcpServer`
- macOS Intel: `publish/macos-x64/AsposeMcpServer`
- macOS ARM: `publish/macos-arm64/AsposeMcpServer`

**注意：** 構建產物為自包含單一可執行檔案，無需安裝 .NET Runtime 即可運行。

### 運行測試

本專案包含完整的單元測試套件，使用 xUnit 測試框架。推薦使用 `test.ps1` 腳本運行測試，它提供了 UTF-8 編碼支援和便捷的參數選項。

**測試統計：**
- **測試類**: 100+ 個測試類（含 Session、Security、Helpers 測試）
- **測試用例**: 7,000+ 個測試用例
- **測試框架**: xUnit 2.9.2

**運行測試：**
```powershell
# 運行所有測試
pwsh test.ps1

# 運行測試（詳細輸出）
pwsh test.ps1 -Verbose

# 運行測試（不重新構建）
pwsh test.ps1 -NoBuild

# 運行測試並收集覆蓋率
pwsh test.ps1 -Coverage

# 運行特定類別的測試
pwsh test.ps1 -Filter "FullyQualifiedName~Word"
pwsh test.ps1 -Filter "FullyQualifiedName~Excel"
pwsh test.ps1 -Filter "FullyQualifiedName~PowerPoint"
pwsh test.ps1 -Filter "FullyQualifiedName~Pdf"

# 運行特定測試類
pwsh test.ps1 -Filter "FullyQualifiedName~WordTextToolTests"

# 運行特定測試方法
pwsh test.ps1 -Filter "FullyQualifiedName~AddTextWithStyle_ShouldCreateEmptyParagraphsWithNormalStyle"

# 跳過授權（強制評估模式）
pwsh test.ps1 -SkipLicense

# 組合使用
pwsh test.ps1 -Verbose -Coverage -Filter "FullyQualifiedName~Word"
```

**test.ps1 參數說明：**
- `-Verbose` - 顯示詳細測試輸出
- `-NoBuild` - 跳過構建步驟（使用已構建的版本）
- `-Coverage` - 收集測試覆蓋率數據
- `-Filter <filter>` - 過濾特定測試（支援 dotnet test 的過濾語法）
- `-SkipLicense` - 跳過授權載入，強制使用評估模式

**測試結構：**
- `Tests/Core/` - 核心功能測試（Handlers、Security、Session、Tracking）
- `Tests/Handlers/` - Handler 測試
  - `Word/` - Word Handler 測試
  - `Excel/` - Excel Handler 測試
  - `PowerPoint/` - PowerPoint Handler 測試
  - `Pdf/` - PDF Handler 測試
- `Tests/Tools/Word/` - Word 工具測試（24 個測試類）
- `Tests/Tools/Excel/` - Excel 工具測試（25 個測試類）
- `Tests/Tools/PowerPoint/` - PowerPoint 工具測試（21 個測試類）
- `Tests/Tools/Pdf/` - PDF 工具測試（15 個測試類）
- `Tests/Tools/Conversion/` - 轉換工具測試（2 個測試類）
- `Tests/Helpers/` - Helper 單元測試（Excel、Word、PowerPoint、PDF Helper 測試）
- `Tests/Infrastructure/` - 測試基礎設施（TestBase、WordTestBase、ExcelTestBase、PdfTestBase、HandlerTestBase）

**CI/CD 集成：**
- 測試已集成到 GitHub Actions 工作流中
- 每次推送或創建 Pull Request 時會自動運行測試
- 測試在評估模式下運行（無需授權檔案）

**測試注意事項：**
- `test.ps1` 腳本會自動設置 UTF-8 編碼，確保中文輸出正常顯示
- 測試會創建臨時檔案，測試完成後會自動清理
- Aspose 授權檔案不會包含在 Git 倉庫中
- 使用 `-SkipLicense` 參數可在評估模式下運行測試（無需授權檔案）
- 測試檔案會保存在系統臨時目錄中
- 測試結果會保存為 `Tests/TestResults/test-results.trx`（TRX 格式）

### 代碼品質檢查

本專案使用 JetBrains 工具進行代碼品質檢查和格式化。推薦使用 `code-quality.ps1` 腳本運行代碼檢查。

**運行代碼品質檢查：**
```powershell
# 執行 CleanupCode 和 InspectCode（預設）
pwsh code-quality.ps1

# 只執行 CleanupCode（代碼格式化）
pwsh code-quality.ps1 -CleanupCode

# 只執行 InspectCode（代碼檢查）
pwsh code-quality.ps1 -InspectCode

# 執行兩個（明確指定）
pwsh code-quality.ps1 -CleanupCode -InspectCode
```

**code-quality.ps1 參數說明：**
- `-CleanupCode` - 執行 JetBrains CleanupCode（代碼格式化）
- `-InspectCode` - 執行 JetBrains InspectCode（代碼檢查，輸出到 `report.xml`）
- `-Profile <profile>` - 指定 CleanupCode 配置檔（預設：`Built-in: Full Cleanup`）
- `-Exclude <patterns>` - 排除的文件模式（預設：`*.txt`）

**注意事項：**
- `code-quality.ps1` 腳本會自動設置 UTF-8 編碼，確保中文輸出正常顯示
- CleanupCode 會格式化代碼，HTML 文件也會被格式化（但 CSS 確保程式碼區塊不會跑版）
- InspectCode 會生成 `report.xml` 報告文件，可用於分析代碼問題

## 📋 工具列表

### Word 文檔處理 (24 個工具)

**檔案操作 (1)**
- `word_file` - 創建、轉換、合併、拆分、從範本創建（5個操作：create, create_from_template, convert, merge, split）
  - `create_from_template` 使用 LINQ Reporting Engine，支援 `<<[ds.Name]>>`、`<<foreach [item in ds.Items]>>...<</foreach>>` 等語法

**內容編輯 (6)**
- `word_text` - 添加、刪除、替換、搜尋、格式化文字（8個操作：add, delete, replace, search, format, insert_at_position, delete_range, add_with_style）
- `word_paragraph` - 插入、刪除、編輯段落格式（7個操作：insert, delete, edit, get, get_format, copy_format, merge）
- `word_table` - 創建、刪除、獲取表格，插入/刪除行列，合併/拆分單元格（16個操作：create, delete, get, insert_row, delete_row, insert_column, delete_column, merge_cells, split_cell, edit_cell_format, move_table, copy_table, get_structure, set_border, set_column_width, set_row_height）
- `word_image` - 添加、編輯、刪除、替換圖片，提取圖片（6個操作：add, edit, delete, get, replace, extract）
- `word_shape` - 添加線條、文字框、圖表、形狀管理，支援邊框樣式：solid, dash, dot, dashDot, dashDotDot, roundDot（9個操作：add_line, add_textbox, get_textboxes, edit_textbox_content, set_textbox_border, add_chart, add, get, delete）
- `word_list` - 添加、編輯、刪除清單項目，重新編號、轉換為清單（8個操作：add_list, add_item, delete_item, edit_item, set_format, get_format, restart_numbering, convert_to_list）

**格式設定 (4)**
- `word_format` - 獲取/設定 Run 格式，獲取/添加/清除定位點，設定段落邊框（6個操作：get_run_format, set_run_format, get_tab_stops, add_tab_stop, clear_tab_stops, set_paragraph_border）
- `word_style` - 獲取、創建、應用樣式，從其他文檔複製樣式（4個操作：get_styles, create_style, apply_style, copy_styles）
- `word_page` - 設定頁邊距、方向、大小、頁碼、刪除頁面、插入空白頁、添加分頁符（8個操作：set_margins, set_orientation, set_size, set_page_number, set_page_setup, delete_page, insert_blank_page, add_page_break）
- `word_header_footer` - 設定頁首頁尾文字、圖片、線條、定位點（10個操作：set_header_text, set_footer_text, set_header_image, set_footer_image, set_header_line, set_footer_line, set_header_tabs, set_footer_tabs, set_header_footer, get）

**高級功能 (13)**
- `word_bookmark` - 添加、編輯、刪除、獲取書籤，跳轉到書籤（5個操作：add, edit, delete, get, goto）
- `word_hyperlink` - 添加、編輯、刪除、獲取超連結（4個操作：add, edit, delete, get）
- `word_comment` - 添加、刪除、獲取註釋，回覆註釋（4個操作：add, delete, get, reply）
- `word_field` - 插入、編輯、刪除、更新、獲取欄位（11個操作：insert_field, edit_field, delete_field, update_field, update_all, get_fields, get_field_detail, add_form_field, edit_form_field, delete_form_field, get_form_fields）
- `word_note` - 添加、編輯、刪除腳註和尾註（8個操作：add_footnote, add_endnote, delete_footnote, delete_endnote, edit_footnote, edit_endnote, get_footnotes, get_endnotes）
- `word_reference` - 添加目錄、更新目錄、添加索引、添加交叉引用（4個操作：add_table_of_contents, update_table_of_contents, add_index, add_cross_reference）
- `word_properties` - 獲取、設定文檔屬性（2個操作：get, set）
- `word_protection` - 保護、解除保護文檔（2個操作：protect, unprotect）
- `word_revision` - 獲取、接受、拒絕修訂，比較文檔（5個操作：get_revisions, accept_all, reject_all, manage, compare）
- `word_section` - 插入、刪除、獲取節資訊（3個操作：insert, delete, get）
- `word_watermark` - 添加文字/圖片水印、移除水印（3個操作：add, add_image, remove）
- `word_mail_merge` - 郵件合併
- `word_content` - 獲取內容、詳細內容、統計資訊、文檔資訊（4個操作：get_content, get_content_detailed, get_statistics, get_document_info）

### Excel 表格處理 (25 個工具)

**檔案操作 (1)**
- `excel_file_operations` - 創建、轉換、合併工作簿、拆分工作簿（4個操作：create, convert, merge, split）

**工作表操作 (1)**
- `excel_sheet` - 添加、刪除、獲取、重新命名、移動、複製、隱藏工作表（7個操作：add, delete, get, rename, move, copy, hide）

**單元格操作 (2)**
- `excel_cell` - 寫入、編輯、獲取、清空單元格（4個操作：write, edit, get, clear）
- `excel_range` - 寫入、編輯、獲取、清空範圍，複製、移動範圍，複製格式（7個操作：write, edit, get, clear, copy, move, copy_format）

**行列操作 (1)**
- `excel_row_column` - 插入/刪除行/列，插入/刪除單元格（6個操作：insert_row, delete_row, insert_column, delete_column, insert_cells, delete_cells）

**資料操作 (1)**
- `excel_data_operations` - 排序、查找替換、批次寫入、獲取內容、統計資訊、獲取已使用範圍（6個操作：sort, find_replace, batch_write, get_content, get_statistics, get_used_range）

**格式與樣式 (2)**
- `excel_style` - 格式化單元格、獲取格式、複製工作表格式（3個操作：format, get_format, copy_sheet_format）
- `excel_conditional_formatting` - 添加、編輯、刪除、獲取條件格式（4個操作：add, edit, delete, get）

**高級功能 (8)**
- `excel_chart` - 添加、編輯、刪除、獲取圖表，更新圖表資料，設定圖表屬性（6個操作：add, edit, delete, get, update_data, set_properties）
- `excel_formula` - 添加、獲取公式，獲取公式結果，計算公式，設定/獲取陣列公式（6個操作：add, get, get_result, calculate, set_array, get_array）
- `excel_pivot_table` - 添加、編輯、刪除、獲取資料透視表，添加/刪除欄位，重新整理（7個操作：add, edit, delete, get, add_field, delete_field, refresh）
- `excel_data_validation` - 添加、編輯、刪除、獲取資料驗證，設定輸入/錯誤訊息（5個操作：add, edit, delete, get, set_messages）
- `excel_image` - 添加、刪除、獲取、提取圖片（4個操作：add, delete, get, extract）
- `excel_hyperlink` - 添加、編輯、刪除、獲取超連結（4個操作：add, edit, delete, get）
- `excel_comment` - 添加、編輯、刪除、獲取批註（4個操作：add, edit, delete, get）
- `excel_named_range` - 添加、刪除、獲取命名範圍（3個操作：add, delete, get）

**保護與設定 (4)**
- `excel_protect` - 保護、解除保護工作簿/工作表，獲取保護資訊，設定單元格鎖定（4個操作：protect, unprotect, get, set_cell_locked）
- `excel_filter` - 應用、移除自動篩選，篩選資料，獲取篩選狀態（4個操作：apply, remove, filter, get_status）
- `excel_freeze_panes` - 凍結、解凍窗格，獲取凍結狀態（3個操作：freeze, unfreeze, get）
- `excel_merge_cells` - 合併、取消合併單元格，獲取合併單元格資訊（3個操作：merge, unmerge, get）

**外觀與視圖 (3)**
- `excel_view_settings` - 設定工作表視圖（縮放、網格線、標題、零值、背景、標籤顏色、凍結窗格、視窗分割、自動調整欄寬列高、顯示公式）（14個操作：set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width, set_row_height, set_background, set_tab_color, set_all, freeze_panes, split_window, auto_fit_column, auto_fit_row, show_formulas）
- `excel_print_settings` - 設定列印區域、標題行、頁面設定（4個操作：set_print_area, set_print_titles, set_page_setup, set_all）
- `excel_group` - 分組/取消分組行/列（4個操作：group_rows, ungroup_rows, group_columns, ungroup_columns）

**屬性與工具 (2)**
- `excel_properties` - 獲取、設定工作簿/工作表屬性（5個操作：get_workbook_properties, set_workbook_properties, get_sheet_properties, edit_sheet_properties, get_sheet_info）
- `excel_get_cell_address` - 單元格地址格式轉換（A1 ↔ 行列索引）

### PowerPoint 演示文稿處理 (21 個工具)

**檔案操作 (1)**
- `ppt_file_operations` - 創建、轉換、合併演示文稿、拆分演示文稿（4個操作：create, convert, merge, split）

**投影片管理 (1)**
- `ppt_slide` - 添加、刪除、獲取投影片資訊、移動、複製、隱藏投影片（8個操作：add, delete, get_info, move, duplicate, hide, clear, edit）

**內容編輯 (5)**
- `ppt_text` - 添加、編輯、替換文字（3個操作：add, edit, replace）
- `ppt_image` - 添加、編輯、刪除、獲取圖片，匯出投影片為圖片、提取圖片（6個操作：add, edit, delete, get, export_slides, extract）
- `ppt_table` - 添加、編輯、刪除表格，插入/刪除行列（9個操作：add, edit, delete, get_content, insert_row, insert_column, delete_row, delete_column, edit_cell）
- `ppt_chart` - 添加、編輯、刪除、獲取圖表，更新圖表資料（5個操作：add, edit, delete, get_data, update_data）
- `ppt_shape` - 統一形狀管理工具（12個操作：get, get_details, delete, edit, set_format, clear_format, group, ungroup, copy, reorder, align, flip）

**格式設定 (2)**
- `ppt_text_format` - 批次格式化文字
- `ppt_background` - 設定投影片背景（顏色/圖片）（2個操作：set, get）

**高級功能 (8)**
- `ppt_animation` - 添加、編輯、刪除、獲取動畫（4個操作：add, edit, delete, get）
- `ppt_transition` - 設定、刪除、獲取轉場效果（3個操作：set, get, delete）
- `ppt_hyperlink` - 添加、編輯、刪除、獲取超連結（4個操作：add, edit, delete, get）
- `ppt_media` - 添加、刪除音訊/影片，設定播放設定（5個操作：add_audio, delete_audio, add_video, delete_video, set_playback）
- `ppt_smart_art` - 添加、管理 SmartArt 節點（2個操作：add, manage_nodes）
- `ppt_section` - 添加、重新命名、刪除章節（4個操作：add, rename, delete, get）
- `ppt_notes` - 設定、獲取、清空講者備註，設定頁首頁尾（4個操作：set, get, clear, set_header_footer）
- `ppt_layout` - 設定投影片版面配置，批次應用版面配置（6個操作：set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme）

**頁面與設定 (4)**
- `ppt_page_setup` - 設定投影片大小、方向、頁尾、頁碼（4個操作：set_size, set_orientation, set_footer, set_slide_numbering）
- `ppt_handout` - 設定講義頁首頁尾（1個操作：set_header_footer）
- `ppt_data_operations` - 獲取統計資訊、內容、投影片詳情（3個操作：get_statistics, get_content, get_slide_details）
- `ppt_properties` - 獲取、設定文檔屬性（2個操作：get, set）

### PDF 檔案處理 (15 個工具)

**檔案操作 (1)**
- `pdf_file` - 創建、合併、拆分、壓縮、加密、線性化PDF（6個操作：create, merge, split, compress, encrypt, linearize）

**內容添加 (5)**
- `pdf_text` - 添加、編輯文字，提取文字（3個操作：add, edit, extract）
- `pdf_image` - 添加、編輯、刪除圖片，提取圖片（5個操作：add, delete, edit, extract, get）
- `pdf_table` - 添加、編輯表格（2個操作：add, edit）
- `pdf_watermark` - 添加水印
- `pdf_page` - 添加、刪除頁面，旋轉頁面，獲取頁面資訊（5個操作：add, delete, rotate, get_details, get_info）

**書籤與註釋 (2)**
- `pdf_bookmark` - 添加、編輯、刪除、獲取書籤（4個操作：add, delete, edit, get）
- `pdf_annotation` - 添加、編輯、刪除、獲取註釋（4個操作：add, delete, edit, get）

**連結與表單 (2)**
- `pdf_link` - 添加、編輯、刪除、獲取超連結（4個操作：add, delete, edit, get）
- `pdf_form_field` - 添加、編輯、刪除、獲取表單欄位（4個操作：add, delete, edit, get）

**附件與簽名 (2)**
- `pdf_attachment` - 添加、刪除、獲取附件（3個操作：add, delete, get）
- `pdf_signature` - 簽名、刪除簽名、獲取簽名（3個操作：sign, delete, get）

**讀取與屬性 (3)**
- `pdf_info` - 獲取PDF內容和統計資訊（2個操作：get_content, get_statistics）
- `pdf_properties` - 獲取、設定文檔屬性（2個操作：get, set）
- `pdf_redact` - 編輯（塗黑）文字或區域

## ⚠️ 重要說明

### 索引行為說明

**索引在刪除操作後會變化：**
- 當執行刪除操作（如刪除段落、表格、圖片等）後，後續元素的索引會自動調整
- 這是正常行為，因為索引是基於當前文檔狀態的
- **建議**：在執行刪除操作後，重新使用 `get` 操作獲取最新的索引列表

**範例：**
```
1. word_image(operation='get', path='doc.docx')  # 返回圖片索引: 0, 1, 2
2. word_image(operation='delete', path='doc.docx', imageIndex=1)  # 刪除索引1的圖片
3. word_image(operation='get', path='doc.docx')  # 現在返回: 0, 1 (原索引2變成1)
```

**paragraphIndex 參數說明：**
- 有效範圍：`0` 到 `段落總數-1`，或使用 `-1` 表示最後一個段落
- 使用 `get` 操作可以獲取當前文檔的段落總數
- 某些操作（如 `word_hyperlink` 的 `add`）會在指定段落**之後**創建新段落，而不是插入到段落內部
- 刪除段落後，後續段落的索引會自動調整

**參數命名一致性：**
- 為了向後兼容，某些參數支援多種命名方式：
  - `startColumn` / `startCol`
  - `columnIndex` / `colIndex`
  - `tableIndex` / `sourceTableIndex`
  - `text` / `replyText` (用於評論回覆)

## 📝 使用範例

### 從A文檔複製格式到B文檔

**複製段落格式：**
```
1. word_paragraph(path="A.docx", operation="get_format", paragraphIndex=0)
2. 使用返回的格式資訊
3. word_paragraph(path="B.docx", operation="edit", paragraphIndex=0, ...)
```

**複製表格結構：**
```
1. word_table(path="A.docx", operation="get_structure", tableIndex=0)
2. 參考返回的結構資訊
3. word_table(path="B.docx", operation="create", ...) 創建相同結構
```

**複製樣式：**
```
word_style(path="B.docx", operation="copy_styles", sourceDocument="A.docx")
```

## 🌍 跨平台支援

支援以下平台（使用 .NET 8.0 自包含發布）：
- ✅ Windows (x64) - `win-x64`
- ✅ Linux (x64) - `linux-x64`
- ✅ macOS Intel (x64) - `osx-x64`
- ✅ macOS ARM (arm64 - M1/M2/M3) - `osx-arm64`

**技術規格：**
- .NET 8.0 Runtime（自包含，無需額外安裝）
- Aspose.Total 23.10.0（包含 Words、Cells、Slides、Pdf、Email）
- 單一可執行檔案（PublishSingleFile）
- 支援 UTF-8 編碼（完整中文支援）

**獲取方式：**
- 從 [GitHub Releases](https://github.com/xjustloveux/aspose-mcp-server/releases) 下載預編譯版本
- 或使用 `publish.ps1` 腳本本地構建

**注意：** GitHub Actions 會在推送到 main/master 分支時自動構建所有平台版本。

### Linux/macOS 額外需求

部分圖片處理功能依賴 `System.Drawing.Common`，在 Linux/macOS 上需要安裝 `libgdiplus`：

**安裝方式：**
```bash
# Ubuntu/Debian
sudo apt-get install libgdiplus

# CentOS/RHEL
sudo yum install libgdiplus

# macOS
brew install mono-libgdiplus
```

**受影響的功能：**

| 功能 | 工具 | 說明 |
|------|------|------|
| 投影片匯出為圖片 | `ppt_image_operations` | `export_slides` 操作 |
| 投影片縮圖生成 | `ppt_data_operations` | `get_slide_details` 的 `includeThumbnail` 參數 |
| 簡報轉圖片 | `ppt_file_operations` | 轉換為圖片格式 |
| PDF 圖片提取 | `pdf_image` | `extract` 操作的某些格式 |

**未安裝 libgdiplus 時的錯誤訊息：**
```
System.TypeInitializationException: The type initializer for 'Gdip' threw an exception.
---> System.PlatformNotSupportedException: System.Drawing.Common is not supported on non-Windows platforms.
```

## 📄 授權

**本專案源代碼授權：**
- 本專案的源代碼採用 [MIT License](LICENSE) 授權
- 您可以自由使用、修改和分發本專案的源代碼

**使用本專案需要 Aspose 授權：**
本專案需要有效的 Aspose 授權檔案才能正常運行。支援以下授權類型：
- `Aspose.Total.lic` - 總授權（包含所有組件，推薦）
- `Aspose.Words.lic` - Word 組件授權
- `Aspose.Cells.lic` - Excel 組件授權
- `Aspose.Slides.lic` - PowerPoint 組件授權
- `Aspose.Pdf.lic` - PDF 組件授權

**授權檔案配置方式（按優先順序）：**
1. **命令列參數**（最高優先級）：`--license 路徑`
2. **環境變數**：設定 `ASPOSE_LICENSE_PATH` 環境變數
3. **自動搜尋**（預設）：在可執行檔案同一目錄搜尋常見授權檔案名稱

**授權搜尋順序：**
1. `Aspose.Words.lic`、`Aspose.Cells.lic`、`Aspose.Slides.lic`、`Aspose.Pdf.lic`（根據啟用的工具）
2. `Aspose.Total.lic`（最後 fallback）

**試用模式：**
如果找不到授權檔案，系統會以試用模式運行（生成的文檔會有試用版標記）。建議配置有效授權以移除標記。

### 評估模式限制

在無授權（評估模式）下運行時，Aspose 組件會有以下限制：

| 組件 | 限制說明 |
|------|----------|
| **Aspose.Words** | 字型格式（名稱、大小、顏色）可能無法正確應用；文檔會包含評估版水印 |
| **Aspose.Cells** | 多工作表操作受限（如跨工作表存取）；工作簿會包含評估版工作表 |
| **Aspose.Slides** | 文字替換時可能包含水印文字；投影片可能包含評估版標記 |
| **Aspose.PDF** | 集合限制為最多 4 個元素（如頁面數限制）；PDF 會包含評估版水印 |

**注意事項：**
- 大部分基本操作（讀取、創建、簡單編輯）在評估模式下仍可正常運行
- 評估模式適合測試和開發，但不建議用於生產環境
- 使用 `test.ps1 -SkipLicense` 可在評估模式下運行單元測試

**授權版本相容性：**
- 當前使用的 Aspose 版本：23.10.0
- 授權檔案需與使用的 Aspose 版本相容

## ❓ 常見問題

### Q: 如何確認工具是否正常運行？
A: 啟動 MCP 客戶端後，檢查工具列表是否包含 `word_*`、`excel_*` 等工具。如果沒有，請檢查：
1. 配置檔案路徑是否正確
2. 可執行檔案是否有執行權限（Linux/macOS）
3. 授權檔案是否正確配置
4. 查看 MCP 客戶端的錯誤日誌

### Q: 為什麼生成的文檔有試用版標記？
A: 這表示授權檔案未正確載入。請檢查：
1. 授權檔案路徑是否正確
2. 授權檔案是否與 Aspose 版本相容（當前版本：23.10.0）
3. 授權檔案是否有效且未過期

### Q: 可以同時啟用多個工具類型嗎？
A: 可以。使用 `--all` 參數或同時指定多個參數，例如：
```json
"args": ["--word", "--excel", "--pdf"]
```

### Q: 轉換工具何時可用？
A: 
- `convert_to_pdf`：啟用任何文檔工具（`--word`、`--excel`、`--ppt`）時自動可用
- `convert_document`：啟用兩個或以上文檔工具時自動可用

### Q: 支援哪些文檔格式？
A: 
- **Word**: DOC、DOCX、RTF、ODT、HTML、TXT 等
- **Excel**: XLS、XLSX、CSV、ODS、HTML 等
- **PowerPoint**: PPT、PPTX、ODP、HTML 等
- **PDF**: PDF（讀寫、編輯、簽名等）

### Q: 如何在 Linux/macOS 上設置執行權限？
A: 
```bash
chmod +x AsposeMcpServer
```

### Q: 錯誤訊息顯示路徑無效怎麼辦？
A: 檢查：
1. 路徑是否使用正確的分隔符（Windows 可用 `/` 或 `\\`）
2. 路徑長度是否超過 260 字元（Windows 限制）
3. 檔案名稱是否包含非法字元
4. 是否啟用了絕對路徑（某些工具可能需要）

### Q: 如何查看詳細的錯誤資訊？
A: 檢查 MCP 客戶端的錯誤日誌。生產環境中，詳細錯誤資訊會被清理以防止資訊洩露。開發環境（DEBUG 模式）會顯示完整錯誤資訊。

### Q: 可以自訂工具嗎？
A: 可以。工具和 Handler 都支援自動發現：

**工具 (Tools)**：
1. 創建新的工具類並添加 `[McpServerToolType]` 屬性
2. 遵循命名約定（`*Tool.cs`）
3. 放在對應的 `Tools/` 子目錄中
4. 工具會透過 MCP SDK 自動註冊

**Handler**：
1. 創建新的 Handler 類並繼承 `OperationHandlerBase<TContext>`
2. 實作 `Operation` 屬性和 `Execute` 方法
3. 放在對應的 `Handlers/{DocumentType}/{Category}/` 目錄中
4. Handler 會透過命名空間自動發現並註冊
5. 如需排除特定 Handler，可標記 `[ExcludeFromAutoDiscovery]` 屬性

## 🔗 相關資源

**官方文檔：**
- [Aspose.Total for .NET](https://products.aspose.com/total/net/)
- [Model Context Protocol](https://modelcontextprotocol.io/) - MCP 官方網站與規範
- [.NET MCP SDK](https://github.com/modelcontextprotocol/csharp-sdk) - 官方 C# SDK（本專案使用）

**MCP 客戶端：**
- [Claude Desktop](https://claude.ai/download) - Anthropic 官方 MCP 客戶端
- [Cursor](https://cursor.sh/) - AI 程式碼編輯器，支援 MCP
- [Continue](https://continue.dev/) - VS Code 擴展，支援 MCP

**專案資源：**
- [GitHub Repository](https://github.com/xjustloveux/aspose-mcp-server)
- [GitHub Pages 文檔](https://xjustloveux.github.io/aspose-mcp-server/) - 完整文檔網站
- [配置範例](config_example.json) - 詳細的 MCP 客戶端配置範例
