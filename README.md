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

**格式支援：** ![Word](https://img.shields.io/badge/Word-DOC%2CDOCX-blue) ![Excel](https://img.shields.io/badge/Excel-XLS%2CXLSX-green) ![PowerPoint](https://img.shields.io/badge/PowerPoint-PPT%2CPPTX-orange) ![PDF](https://img.shields.io/badge/PDF-PDF-red) ![OCR](https://img.shields.io/badge/OCR-PNG%2CJPG%2CPDF-purple) ![Email](https://img.shields.io/badge/Email-EML%2CMSG-teal) ![BarCode](https://img.shields.io/badge/BarCode-QR%2CCode128-brown)

基於 .NET 8.0 和 Aspose.Total 的 Model Context Protocol (MCP) 伺服器，為 MCP 客戶端提供強大的辦公文檔處理能力。

## ✨ 特性

### 核心功能
- **115 個統一工具** - Word(27)、Excel(31)、PowerPoint(25)、PDF(19)、OCR(2)、Email(6)、BarCode(2)、轉換(2)、Session(1) 已整合
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

## 🚀 快速開始

### 1. 下載

**方法 A：從 GitHub Releases 下載**

從 [GitHub Releases](https://github.com/xjustloveux/aspose-mcp-server/releases) 下載對應平台版本：

| 平台 | 檔案 |
|------|------|
| Windows | `aspose-mcp-server-windows-x64.zip` |
| Linux | `aspose-mcp-server-linux-x64.tar.gz` |
| macOS Intel | `aspose-mcp-server-macos-x64.tar.gz` |
| macOS ARM | `aspose-mcp-server-macos-arm64.tar.gz` |

**方法 B：macOS 使用 Homebrew 安裝（推薦）**

```bash
brew install xjustloveux/tap/aspose-mcp-server
```

### 2. 配置 MCP 客戶端

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

**可用參數**（不帶任何工具參數時，預設啟用所有工具）:
- `--word` - Word 工具（自動包含轉換功能）
- `--excel` - Excel 工具（自動包含轉換功能）
- `--powerpoint` / `--ppt` - PowerPoint 工具（自動包含轉換功能）
- `--pdf` - PDF 工具
- `--ocr` - OCR 文字辨識工具
- `--email` - Email 工具
- `--barcode` - BarCode 工具
- `--all` - 所有工具（等同不帶工具參數）
- `--session-enabled` - 啟用 Session 管理（`document_session` 工具）
- `--license 路徑` - 指定授權檔案路徑（可選）

> **工具過濾**：指定工具參數時，只有啟用的工具類別會出現在 MCP 工具列表中。例如使用 `--word` 時，只會顯示 `word_*` 相關工具。

**轉換功能說明**：
- 啟用任何文檔工具（`--word`、`--excel`、`--ppt`）或 `--pdf` 時，自動包含 `convert_to_pdf`（轉換為PDF）
- 啟用兩個或以上文檔工具時，自動包含 `convert_document`（跨格式轉換，如Word轉Excel）

📋 **更多配置範例：** `config_example.json`（配置格式適用於所有 MCP 客戶端）

### 3. 重啟 MCP 客戶端

完成配置後，重啟您使用的 MCP 客戶端（如 Claude Desktop、Cursor 等）即可開始使用。

## 📦 功能概覽

| 模組 | 工具數 | 主要功能 |
|------|--------|---------|
| **Word** | 27 | 檔案操作、文字/段落/表格/圖片編輯、格式/樣式/頁面設定、書籤/超連結/註釋/目錄/修訂/郵件合併/數位簽章/內容控制項/渲染 |
| **Excel** | 31 | 檔案/工作表/行列/單元格操作、排序/篩選/驗證、圖表/公式/樞紐分析、表格/形狀/迷你圖/JSON匯入/渲染 |
| **PowerPoint** | 25 | 檔案/投影片管理、文字/圖片/表格/圖表/形狀/SmartArt/媒體、動畫/轉場/備註/註解/加密/浮水印/字型管理 |
| **PDF** | 19 | 檔案操作（含加密/解密）、文字/圖片/表格/水印/頁面（含裁切）、書籤/註釋/表單（含匯入匯出）、頁首頁尾/印章/目錄/PDF/A合規 |
| **OCR** | 2 | 影像前處理（校正/降噪/對比/縮放）、文字辨識（圖片/PDF/收據/身分證/護照） |
| **Email** | 6 | 郵件建立/讀取/轉換、內容編輯、附件管理、日曆事件、聯絡人、格式轉換（EML↔MSG↔HTML） |
| **BarCode** | 2 | 條碼產生（QR/Code128/EAN13 等）、條碼辨識（自動偵測/指定類型） |
| **轉換** | 2 | `convert_to_pdf`（Word/Excel/PPT/HTML/EPUB/Markdown/SVG→PDF）、`convert_document`（跨格式轉換） |

> 📖 完整工具列表與操作說明請參閱 [工具列表](https://xjustloveux.github.io/aspose-mcp-server/tools.html)

## 🔌 傳輸模式

| 模式 | 命令 | 端點 | 適用場景 |
|------|------|------|---------|
| **Stdio**（預設） | `AsposeMcpServer.exe --word` | - | 本地 MCP 客戶端 |
| **HTTP** | `AsposeMcpServer.exe --http --port 3000 --word` | `http://localhost:3000/mcp` | 網頁應用 |
| **WebSocket** | `AsposeMcpServer.exe --ws --port 3000 --word` | `ws://localhost:3000/ws` | 即時互動 |

**環境變數：**

| 變數 | 說明 | 預設值 |
|------|------|--------|
| `ASPOSE_TRANSPORT` | 傳輸模式 (stdio/http/ws) | stdio |
| `ASPOSE_PORT` | 監聽埠號 | 3000 |
| `ASPOSE_HOST` | 監聽位址（`localhost`、`0.0.0.0`、`*`） | localhost |
| `ASPOSE_TOOLS` | 啟用的工具（all 或 word,excel,pdf,ppt,ocr,email,barcode） | 全部啟用 |

> **注意**: Docker/Kubernetes 部署時需設定 `ASPOSE_HOST=0.0.0.0` 以便容器外部可以訪問。

## 🔒 安全特性

- **路徑驗證** - 所有檔案路徑經 `SecurityHelper.ValidateFilePath()` 驗證，防止路徑遍歷攻擊
- **輸入驗證** - 陣列大小上限 1000 項、字串長度上限 10000 字元
- **錯誤處理** - 錯誤訊息清理，防止資訊洩露（移除路徑、堆疊追蹤等敏感資訊）
- **Origin 驗證** - HTTP/WebSocket 模式預設啟用，防止 DNS 重綁定攻擊

| 限制項目 | 上限值 |
|---------|--------|
| 最大路徑長度 | 260 字元 |
| 最大檔案名稱長度 | 255 字元 |
| 最大陣列大小 | 1000 項 |
| 最大字串長度 | 10000 字元 |

> 📖 完整安全配置請參閱 [功能特性](https://xjustloveux.github.io/aspose-mcp-server/features.html)

## 🌍 跨平台支援

| 平台 | 文檔處理 | OCR | 備註 |
|------|---------|-----|------|
| Windows x64 | ✅ | ✅ | |
| Linux x64 | ✅ | ✅ | 不需要 libgdiplus |
| macOS Intel x64 | ✅ | ✅ | |
| macOS ARM64 (M1/M2/M3) | ✅ | ✅ | PPT/OCR 需 Rosetta 2 |
| Linux ARM64 | ✅ | ❌ | ONNX Runtime 限制 |

**跨平台方案：** PowerPoint 使用 `Aspose.Slides.NET6.CrossPlatform`、PDF 使用 `Aspose.PDF.Drawing`、Word/Excel 使用 `SkiaSharp`，全部無需外部圖形庫。

**Linux 字型：** 建議安裝 `fonts-liberation`（英文）和 `fonts-noto-cjk`（中日韓文），Docker 映像已內建基本 CJK 字型。

> 📖 詳細平台需求與字型配置請參閱 [部署指南](https://xjustloveux.github.io/aspose-mcp-server/deployment.html)

## 📄 授權

**本專案源代碼**採用 [MIT License](LICENSE) 授權。**運行時需要 Aspose 授權**，支援：

- `Aspose.Total.lic` - 總授權（包含所有組件，推薦）
- 單一組件授權：`Aspose.Words.lic`、`Aspose.Cells.lic`、`Aspose.Slides.lic`、`Aspose.Pdf.lic`、`Aspose.OCR.lic`、`Aspose.Email.lic`、`Aspose.BarCode.lic`

**配置方式（按優先順序）：**
1. **命令列參數**：`--license 路徑`
2. **環境變數**：`ASPOSE_LICENSE_PATH`
3. **自動搜尋**（預設）：在可執行檔案同一目錄搜尋

找不到授權時以試用模式運行（文檔含評估版標記）。使用 `test.ps1 -SkipLicense` 可在評估模式下運行測試。

## ⚠️ 重要說明

### 索引行為說明

**索引在刪除操作後會變化：**
- 當執行刪除操作（如刪除段落、表格、圖片等）後，後續元素的索引會自動調整
- **建議**：在執行刪除操作後，重新使用 `get` 操作獲取最新的索引列表

```
1. word_image(operation='get', path='doc.docx')  # 返回圖片索引: 0, 1, 2
2. word_image(operation='delete', path='doc.docx', imageIndex=1)  # 刪除索引1的圖片
3. word_image(operation='get', path='doc.docx')  # 現在返回: 0, 1 (原索引2變成1)
```

**paragraphIndex 參數說明：**
- 有效範圍：`0` 到 `段落總數-1`，或使用 `-1` 表示最後一個段落
- 某些操作會在指定段落**之後**創建新段落，而不是插入到段落內部

**參數命名一致性：** 為向後兼容，某些參數支援多種命名（如 `startColumn` / `startCol`、`columnIndex` / `colIndex`）。

## 📝 使用範例

### 從A文檔複製格式到B文檔

**複製段落格式：**
```
1. word_paragraph(path="A.docx", operation="get_format", paragraphIndex=0)
2. 使用返回的格式資訊
3. word_paragraph(path="B.docx", operation="edit", paragraphIndex=0, ...)
```

**複製樣式：**
```
word_style(path="B.docx", operation="copy_styles", sourceDocument="A.docx")
```

## 🔗 相關資源

| 類別 | 連結 |
|------|------|
| **完整文檔** | [GitHub Pages](https://xjustloveux.github.io/aspose-mcp-server/) — 功能特性、工具列表、快速開始、開發者指南、部署指南、FAQ |
| **配置範例** | [config_example.json](config_example.json) |
| **Aspose** | [Aspose.Total for .NET](https://products.aspose.com/total/net/) |
| **MCP** | [MCP 官方網站](https://modelcontextprotocol.io/) · [.NET MCP SDK](https://github.com/modelcontextprotocol/csharp-sdk) |
| **MCP 客戶端** | [Claude Desktop](https://claude.ai/download) · [Cursor](https://cursor.sh/) · [Continue](https://continue.dev/) |
| **專案** | [GitHub Repository](https://github.com/xjustloveux/aspose-mcp-server) |

> 📖 進階主題（Session 管理、認證機制、追蹤系統、部署指南、開發者指南、常見問答）請參閱 [完整文檔](https://xjustloveux.github.io/aspose-mcp-server/)。
