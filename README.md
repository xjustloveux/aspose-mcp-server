# Aspose MCP Server

基於 .NET 8.0 和 Aspose.Total 的 Model Context Protocol (MCP) 伺服器，為 AI 助手提供強大的辦公文檔處理能力。

## ✨ 特性

- **90 個統一工具** - Word(24)、Excel(25)、PowerPoint(24)、PDF(15)、轉換工具(2)已整合
- **按需啟用** - 只啟用需要的文檔類型
- **跨平台** - Windows、Linux、macOS (Intel + ARM)
- **開箱即用** - publish/ 包含預編譯版本
- **完整讀寫** - 支援從A文檔讀取格式應用到B文檔
- **安全加固** - 全面的路徑驗證、輸入驗證和錯誤處理

## 🚀 快速開始

### 1. 下載預編譯版本

從 [GitHub Releases](../../releases) 下載最新版本：
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
3. **命令列參數**：使用 `--license:路徑` 或 `--license=路徑` 指定授權檔案

**範例：**
```json
{
  "mcpServers": {
    "aspose-word": {
      "command": "C:/Tools/aspose-mcp-server/AsposeMcpServer.exe",
      "args": ["--word", "--license:C:/Licenses/Aspose.Words.lic"]
    }
  }
}
```

**注意**：如果找不到授權檔案，系統會以試用模式運行（會有試用版標記）。

### 2. 配置 Claude Desktop

編輯配置檔案：
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

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
- `--license:路徑` 或 `--license=路徑` - 指定授權檔案路徑（可選）

**轉換功能說明**：
- 啟用任何文檔工具（`--word`、`--excel`、`--ppt`）時，自動包含 `convert_to_pdf`（轉換為PDF）
- 啟用兩個或以上文檔工具時，自動包含 `convert_document`（跨格式轉換，如Word轉Excel）

📋 **更多配置範例：** `claude_desktop_config_example.json`

### 3. 重啟 Claude Desktop

完成！

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

### PowerPoint (24個工具)

**檔案操作**：創建、讀取、轉換、合併、拆分  
**投影片管理**：添加、刪除、移動、複製、隱藏、設定版面配置、設定大小  
**內容編輯**：文字、圖片、表格、圖表、形狀、SmartArt、媒體（音訊/影片）  
**格式設定**：文字格式、形狀格式、背景、頁眉頁腳、主題  
**高級功能**：動畫、轉場、備註、章節、超連結、文檔屬性、保護

### PDF (15個工具)

**檔案操作**：創建、讀取、合併、拆分、壓縮、加密  
**內容添加**：文字、圖片、表格、水印、頁面、書籤、註釋、連結、表單欄位、附件  
**編輯操作**：編輯文字、表格、書籤、註釋、連結、表單欄位、圖片  
**讀取操作**：提取文字、圖片、讀取頁面資訊、書籤、註釋、連結、表單欄位、附件、簽名、統計資訊  
**高級功能**：簽名、頁面旋轉、編輯（塗黑）

### 轉換工具 (2個)

- `convert_to_pdf` - 將任何文檔轉換為PDF（啟用任何文檔工具時自動可用）
- `convert_document` - 跨格式轉換（啟用兩個或以上文檔工具時自動可用）

## 🔒 安全特性

### 路徑驗證
- ✅ 所有檔案路徑都經過 `SecurityHelper.ValidateFilePath()` 驗證
- ✅ 防止路徑遍歷攻擊（`../`, `..\`）
- ✅ 限制路徑長度（最大260字元）和檔案名稱長度（最大255字元）
- ✅ 驗證路徑中的非法字元

### 輸入驗證
- ✅ 陣列大小驗證（`SecurityHelper.ValidateArraySize`，最大1000項）
- ✅ 字串長度驗證（`SecurityHelper.ValidateStringLength`，最大10000字元）

### 錯誤處理
- ✅ 錯誤訊息清理（`McpErrorHandler.SanitizeErrorMessage`），防止資訊洩露
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
├── Tools/            📁 工具原始碼
│   ├── Word/         24 個工具
│   ├── Excel/        25 個工具
│   ├── PowerPoint/   24 個工具
│   ├── PDF/          15 個工具
│   └── Conversion/   2 個工具
├── Core/             🔧 MCP 伺服器核心
│   ├── SecurityHelper.cs      - 安全驗證工具
│   ├── McpErrorHandler.cs     - 錯誤處理
│   ├── ToolRegistry.cs        - 工具註冊
│   └── ServerConfig.cs        - 伺服器配置
├── .github/
│   └── workflows/    🔄 GitHub Actions 工作流程
└── bin/              ❌ 本地編譯輸出（不在版本控制）
```

### 本地開發

```bash
# 複製倉庫
git clone <repository-url>
cd aspose-mcp-server

# 編譯 Release 版本
pwsh build.ps1 --configuration Release

# 發布 Windows 版本
pwsh publish.ps1 -Windows

# 發布所有平台
pwsh publish.ps1 -All
```

### 多平台構建

**所有平台由 GitHub Actions 自動構建：**
- 推送到 main/master 分支時自動觸發
- 構建產物：從 Actions 頁面或 Releases 頁面下載

## 📋 工具列表

### Word 文檔處理 (24 個工具)

**檔案操作 (1)**
- `word_file` - 創建、讀取、轉換、合併、拆分、從範本創建（5個操作：create, create_from_template, convert, merge, split）

**內容編輯 (6)**
- `word_text` - 添加、刪除、替換、搜尋、格式化文字（8個操作：add, delete, replace, search, format, insert_at_position, delete_range, add_with_style）
- `word_paragraph` - 插入、刪除、編輯段落格式（7個操作：insert, delete, edit, get, get_format, copy_format, merge）
- `word_table` - 添加、編輯、刪除表格，插入/刪除行列，合併/拆分單元格（17個操作：add_table, edit_table_format, delete_table, get_tables, insert_row, delete_row, insert_column, delete_column, merge_cells, split_cell, edit_cell_format, move_table, copy_table, get_table_structure, set_table_border, set_column_width, set_row_height）
- `word_image` - 添加、編輯、刪除、替換圖片，提取圖片（6個操作：add, edit, delete, get, replace, extract）
- `word_shape` - 添加線條、文字框、圖表（6個操作：add_line, add_textbox, get_textboxes, edit_textbox_content, set_textbox_border, add_chart）
- `word_list` - 添加、編輯、刪除清單項目（6個操作：add_list, add_item, delete_item, edit_item, set_format, get_format）

**格式設定 (4)**
- `word_format` - 獲取/設定 Run 格式，獲取定位點，設定段落邊框（4個操作：get_run_format, set_run_format, get_tab_stops, set_paragraph_border）
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
- `word_watermark` - 添加水印（1個操作：add）
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
- `excel_image` - 添加、刪除、獲取圖片（3個操作：add, delete, get）
- `excel_hyperlink` - 添加、編輯、刪除、獲取超連結（4個操作：add, edit, delete, get）
- `excel_comment` - 添加、編輯、刪除、獲取批註（4個操作：add, edit, delete, get）
- `excel_named_range` - 添加、刪除、獲取命名範圍（3個操作：add, delete, get）

**保護與設定 (4)**
- `excel_protect` - 保護、解除保護工作簿/工作表，獲取保護資訊，設定單元格鎖定（4個操作：protect, unprotect, get, set_cell_locked）
- `excel_filter` - 應用、移除自動篩選，獲取篩選狀態（3個操作：apply, remove, get_status）
- `excel_freeze_panes` - 凍結、解凍窗格，獲取凍結狀態（3個操作：freeze, unfreeze, get）
- `excel_merge_cells` - 合併、取消合併單元格，獲取合併單元格資訊（3個操作：merge, unmerge, get）

**外觀與視圖 (3)**
- `excel_view_settings` - 設定工作表視圖（縮放、網格線、標題、零值、背景、標籤顏色、視窗分割）（10個操作：set_zoom, set_gridlines, set_headers, set_zero_values, set_column_width, set_row_height, set_background, set_tab_color, set_all, split_window）
- `excel_print_settings` - 設定列印區域、標題行、頁面設定（4個操作：set_print_area, set_print_titles, set_page_setup, set_all）
- `excel_group` - 分組/取消分組行/列（4個操作：group_rows, ungroup_rows, group_columns, ungroup_columns）

**屬性與工具 (2)**
- `excel_properties` - 獲取、設定工作簿/工作表屬性（5個操作：get_workbook_properties, set_workbook_properties, get_sheet_properties, edit_sheet_properties, get_sheet_info）
- `excel_get_cell_address` - 單元格地址格式轉換（A1 ↔ 行列索引）

### PowerPoint 演示文稿處理 (24 個工具)

**檔案操作 (1)**
- `ppt_file_operations` - 創建、轉換、合併演示文稿、拆分演示文稿（4個操作：create, convert, merge, split）

**投影片管理 (1)**
- `ppt_slide` - 添加、刪除、獲取投影片資訊、移動、複製、隱藏投影片（8個操作：add, delete, get_info, move, duplicate, hide, clear, edit）

**內容編輯 (5)**
- `ppt_text` - 添加、編輯、替換文字（3個操作：add, edit, replace）
- `ppt_image` - 添加、編輯、刪除圖片（2個操作：add, edit）
- `ppt_table` - 添加、編輯、刪除表格，插入/刪除行列（9個操作：add, edit, delete, get_content, insert_row, insert_column, delete_row, delete_column, edit_cell）
- `ppt_chart` - 添加、編輯、刪除、獲取圖表，更新圖表資料（5個操作：add, edit, delete, get_data, update_data）
- `ppt_shape` - 添加、編輯、刪除、獲取形狀，設定形狀格式（4個操作：edit, delete, get, get_details）

**格式設定 (4)**
- `ppt_text_format` - 批次格式化文字
- `ppt_shape_format` - 設定形狀位置、尺寸、旋轉、填充、線條（2個操作：set, get）
- `ppt_background` - 設定投影片背景（顏色/圖片）（2個操作：set, get）
- `ppt_header_footer` - 設定頁眉頁尾、頁碼、日期（4個操作：set_header, set_footer, batch_set, set_slide_numbering）

**高級功能 (8)**
- `ppt_animation` - 添加、編輯、刪除動畫（3個操作：add, edit, delete）
- `ppt_transition` - 設定、刪除、獲取轉場效果（3個操作：set, get, delete）
- `ppt_hyperlink` - 添加、編輯、刪除、獲取超連結（4個操作：add, edit, delete, get）
- `ppt_media` - 添加、刪除音訊/影片，設定播放設定（5個操作：add_audio, delete_audio, add_video, delete_video, set_playback）
- `ppt_smart_art` - 添加、管理 SmartArt 節點（2個操作：add, manage_nodes）
- `ppt_section` - 添加、重新命名、刪除章節（4個操作：add, rename, delete, get）
- `ppt_notes` - 添加、編輯、獲取、清空講者備註（4個操作：add, edit, get, clear）
- `ppt_layout` - 設定投影片版面配置，批次應用版面配置（6個操作：set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme）

**操作與設定 (5)**
- `ppt_shape_operations` - 對齊形狀、調整順序、組合/取消組合、翻轉形狀、複製形狀（6個操作：group, ungroup, copy, reorder, align, flip）
- `ppt_image_operations` - 替換圖片、提取圖片、匯出投影片為圖片（3個操作：export_slides, extract_images, replace_with_compression）
- `ppt_data_operations` - 批次替換文字、批次設定頁眉頁尾（3個操作：get_statistics, get_content, get_slide_details）
- `ppt_slide_settings` - 設定投影片大小、方向、編號（2個操作：set_size, set_orientation）
- `ppt_properties` - 獲取、設定文檔屬性（2個操作：get, set）

### PDF 檔案處理 (15 個工具)

**檔案操作 (1)**
- `pdf_file` - 創建、合併、拆分、壓縮、加密PDF（5個操作：create, merge, split, compress, encrypt）

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

## 🎉 主要特性

### MCP 2025-11-25 規範支援
- ✅ 符合最新 MCP 協議規範（protocolVersion: 2025-11-25）
- ✅ 自動工具註解（readonly/destructive）基於命名約定
- ✅ 完整的 JSON-RPC 2.0 錯誤處理

### 統一字型設定
多個工具支援中英文字型分別設定（`fontNameAscii` 和 `fontNameFarEast` 參數）

### 靈活的授權配置
- 支援總授權或單一組件授權
- 自動搜尋、環境變數或命令列參數配置
- 試用模式降級（找不到授權時）

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
1. word_table(path="A.docx", operation="get_table_structure", tableIndex=0)
2. 參考返回的結構資訊
3. word_table(path="B.docx", operation="add_table", ...) 創建相同結構
```

**複製樣式：**
```
word_style(path="B.docx", operation="copy_styles", sourceDocument="A.docx")
```

## 🌍 跨平台支援

所有平台由 **GitHub Actions** 自動構建和發布：
- ✅ Windows (x64)
- ✅ Linux (x64)
- ✅ macOS Intel (x64)
- ✅ macOS ARM (arm64 - M1/M2/M3)

**獲取方式：** 從 [GitHub Releases](../../releases) 下載最新版本

## 📄 授權

本專案需要有效的 Aspose 授權檔案。支援以下授權類型：
- `Aspose.Total.lic` - 總授權（包含所有組件）
- `Aspose.Words.lic`、`Aspose.Cells.lic`、`Aspose.Slides.lic`、`Aspose.Pdf.lic` - 單一組件授權

**配置方式：**
1. 將授權檔案放在可執行檔案同一目錄（自動搜尋）
2. 使用環境變數 `ASPOSE_LICENSE_PATH` 指定路徑
3. 使用命令列參數 `--license:路徑` 指定路徑

如果找不到授權檔案，系統會以試用模式運行（會有試用版標記）。

## 🔗 相關資源

- [Aspose.Total for .NET](https://products.aspose.com/total/net/)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [Claude Desktop](https://claude.ai/desktop)
