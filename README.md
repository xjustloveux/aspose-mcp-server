# Aspose MCP Server

基於 .NET 8.0 和 Aspose.Total 的 Model Context Protocol (MCP) 服務器，為 AI 助手提供強大的辦公文檔處理能力。

## ✨ 特性

- **400+ 個工具** - Word(137)、Excel(121)、PowerPoint(97)、PDF(47)、轉換工具已集成
- **按需啟用** - 只啟用需要的文檔類型
- **跨平台** - Windows、Linux、macOS (Intel + ARM)
- **開箱即用** - publish/ 包含預編譯版本
- **完整讀寫** - 支援從A文檔讀取格式應用到B文檔

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

**放置授權文件：** 將授權文件放在可執行文件同一目錄。支援以下方式：

- **總授權**：`Aspose.Total.lic`（包含所有組件）
- **單一組件授權**：`Aspose.Words.lic`、`Aspose.Cells.lic`、`Aspose.Slides.lic`、`Aspose.Pdf.lic`
- **自訂檔名**：可透過環境變數或命令列參數指定

**授權文件配置方式：**

1. **自動搜尋**（推薦）：將授權文件放在可執行文件目錄，系統會自動搜尋
2. **環境變數**：設定 `ASPOSE_LICENSE_PATH` 環境變數指向授權文件路徑
3. **命令列參數**：使用 `--license:路徑` 或 `--license=路徑` 指定授權文件

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

**注意**：如果找不到授權文件，系統會以試用模式運行（會有試用版標記）。

### 2. 配置 Claude Desktop

編輯配置文件：
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
- `--license:路徑` 或 `--license=路徑` - 指定授權文件路徑（可選）

**轉換功能說明**：
- 啟用任何文檔工具（`--word`、`--excel`、`--ppt`）時，自動包含 `convert_to_pdf`（轉換為PDF）
- 啟用兩個或以上文檔工具時，自動包含 `convert_document`（跨格式轉換，如Word轉Excel）

📋 **更多配置範例：** `claude_desktop_config_example.json`

### 3. 重啟 Claude Desktop

完成！

## 📦 功能概覽

### Word (137個)
**基本操作**：創建、讀取、轉換、合併、拆分、搜尋、統計  
**內容編輯**：文字（中英文字型分別設定）、表格、圖片、圖表、清單、文字框、欄位  
**格式設定**：段落、字型、樣式（複製樣式保留字型）、頁首頁尾、頁面設定  
**高級功能**：書籤、超連結、註解、目錄、文檔屬性、保護、郵件合併、腳注、尾注、交叉引用、索引、文檔比較、修訂管理、表單欄位

### Excel (121個)
**基本操作**：創建、讀取、寫入、轉換、保護  
**工作表操作**：添加、刪除、重命名、複製、移動、隱藏/顯示、讀取資訊  
**行列操作**：插入/刪除行/列、設定行高列寬  
**單元格操作**：合併/取消合併、插入/刪除單元格、鎖定/解鎖  
**數據操作**：排序、篩選、數據驗證  
**格式設定**：單元格格式、條件格式  
**高級功能**：圖表、公式、樞紐表、凍結窗格、超連結、圖片、頁面設定、陣列公式、列印設定、工作表外觀設定

### PowerPoint (97個)
投影片管理、文字、圖片、表格、圖表、動畫、主題、備註、背景、轉場、媒體、批量替換/匯出、編輯操作、讀取操作、刪除操作、文檔操作、形狀操作

### PDF (47個)
創建、讀取、合併、拆分、文字、圖片、表格、水印、加密、簽章、書籤、註解、編輯操作、讀取操作、刪除操作、頁面操作、連結、表單欄位、文檔屬性、壓縮

## 🛠️ 開發者指南

### 倉庫結構
```
aspose-mcp-server/
├── Tools/            📁 工具源代碼
│   ├── Word/         137 個工具
│   ├── Excel/        121 個工具
│   ├── PowerPoint/   100 個工具
│   └── PDF/          47 個工具
├── Core/             🔧 MCP 服務器核心
├── .github/
│   └── workflows/    🔄 GitHub Actions 工作流
└── bin/              ❌ 本地編譯輸出（不在版本控制）
```

### 本地開發

```bash
# 克隆倉庫
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

### Word 文件處理 (137 個)

**基本操作 (7)**
- `word_create` - 創建文檔
- `word_get_content` - 讀取文檔內容
- `word_get_content_detailed` - 讀取詳細內容（包含格式）
- `word_extract_images` - 提取圖片
- `word_convert` - 轉換格式
- `word_merge` - 合併文檔
- `word_split` - 拆分文檔

**內容添加 (15)**
- `word_add_text` - 添加文字
- `word_add_text_with_style` - 添加帶樣式文字（支援縮排、定位點、指定位置插入、中英文字型分別設定）
- `word_add_table` - 添加表格（支援格式、背景色、合併、中英文字型）
- `word_add_image` - 添加圖片（支援大小、對齊、環繞、標題）
- `word_add_chart` - 添加圖表
- `word_add_list` - 添加清單
- `word_add_list_item` - 添加清單項目
- `word_add_line` - 添加線條
- `word_add_textbox` - 添加文字框
- `word_add_page_break` - 添加分頁符
- `word_add_section_break` - 添加分節符
- `word_add_table_of_contents` - 添加目錄
- `word_add_hyperlink` - 添加超連結
- `word_add_bookmark` - 添加書籤
- `word_add_comment` - 添加註解

**編輯操作 (24)**
- `word_edit_paragraph` - 編輯段落格式
- `word_edit_table` - 編輯表格格式
- `word_edit_table_cell` - 編輯表格單元格
- `word_edit_image` - 編輯圖片
- `word_edit_list_item` - 編輯清單項目
- `word_edit_textbox_content` - 編輯文字框內容
- `word_edit_hyperlink` - 編輯超連結
- `word_format_text` - 格式化文字（Run層級）
- `word_replace_text` - 替換文字
- `word_replace_image` - 替換圖片
- `word_insert_paragraph` - 插入段落
- `word_insert_table_row` - 插入表格行
- `word_insert_table_column` - 插入表格列
- `word_insert_blank_page` - 插入空白頁
- `word_insert_field` - 插入功能變數（日期、頁碼等）
- `word_get_fields` - 讀取所有功能變數列表
- `word_get_field_detail` - 讀取功能變數詳細資訊
- `word_update_field` - 更新功能變數
- `word_delete_field` - 刪除功能變數
- `word_merge_paragraphs` - 合併段落
- `word_merge_table_cells` - 合併表格單元格
- `word_split_table_cell` - 拆分表格單元格
- `word_reply_comment` - 回覆註解
- `word_mail_merge` - 郵件合併

**刪除操作 (10)**
- `word_delete_paragraph` - 刪除段落
- `word_delete_table` - 刪除表格
- `word_delete_table_row` - 刪除表格行
- `word_delete_table_column` - 刪除表格列
- `word_delete_image` - 刪除圖片
- `word_delete_list_item` - 刪除清單項目
- `word_delete_text` - 刪除文字
- `word_delete_page` - 刪除頁面
- `word_delete_hyperlink` - 刪除超連結
- `word_delete_bookmark` - 刪除書籤
- `word_delete_comment` - 刪除註解

**格式設定 (14)**
- `word_set_paragraph_border` - 設定段落邊框
- `word_set_table_border` - 設定表格邊框
- `word_set_table_row_height` - 設定表格行高
- `word_set_table_column_width` - 設定表格列寬
- `word_set_textbox_border` - 設定文字框邊框
- `word_set_list_format` - 設定清單格式
- `word_set_page_setup` - 設定頁面
- `word_set_page_number` - 設定頁碼
- `word_set_header_footer` - 設定頁首頁尾（綜合工具）
- `word_set_properties` - 設定文檔屬性
- `word_protect` - 保護文檔
- `word_unprotect` - 解除文檔保護
- `word_manage_revisions` - 接受/拒絕修訂
- `word_add_watermark` - 添加浮水印

**頁首頁尾細粒度控制 (8)**
- `word_set_header_text` - 設定頁首文字
- `word_set_footer_text` - 設定頁尾文字
- `word_set_header_image` - 設定頁首圖片
- `word_set_footer_image` - 設定頁尾圖片
- `word_set_header_line` - 設定頁首線條
- `word_set_footer_line` - 設定頁尾線條
- `word_set_header_tab_stops` - 設定頁首定位點
- `word_set_footer_tab_stops` - 設定頁尾定位點

**讀取與診斷 (9)**
- `word_get_styles` - 讀取樣式
- `word_get_document_info` - 讀取文檔資訊
- `word_get_tab_stops` - 讀取定位點
- `word_get_statistics` - 讀取統計資訊
- `word_get_paragraph_format` - 讀取段落格式
- `word_get_table_structure` - 讀取表格結構
- `word_get_hyperlinks` - 讀取超連結
- `word_get_bookmarks` - 讀取書籤
- `word_get_comments` - 讀取註解

**樣式與格式複製 (4)**
- `word_create_style` - 創建樣式
- `word_copy_styles_from` - 從其他文檔複製樣式
- `word_copy_paragraph_format` - 複製段落格式

**搜尋與導航 (3)**
- `word_search_text` - 搜尋文字（支援正則表達式）
- `word_goto_bookmark` - 跳轉到書籤
- `word_create_from_template` - 從範本創建

**腳注與尾注 (8)**
- `word_add_footnote` - 添加腳注
- `word_add_endnote` - 添加尾注
- `word_get_footnotes` - 讀取所有腳注
- `word_get_endnotes` - 讀取所有尾注
- `word_edit_footnote` - 編輯腳注
- `word_edit_endnote` - 編輯尾注
- `word_delete_footnote` - 刪除腳注
- `word_delete_endnote` - 刪除尾注

**交叉引用與索引 (3)**
- `word_add_cross_reference` - 添加交叉引用（標題、書籤、圖表等）
- `word_add_index` - 添加索引（XE欄位和INDEX欄位）
- `word_update_table_of_contents` - 更新目錄

**樣式應用 (1)**
- `word_apply_style` - 應用樣式到段落、表格或所有段落

**文檔屬性 (2)**
- `word_get_document_properties` - 讀取文檔屬性（元數據）
- `word_set_document_properties` - 設定文檔屬性

**節操作 (3)**
- `word_get_sections_info` - 讀取節資訊
- `word_insert_section` - 插入新節
- `word_delete_section` - 刪除節

**文字操作增強 (3)**
- `word_delete_text_range` - 刪除文字範圍
- `word_insert_text_at_position` - 在指定位置插入文字
- `word_get_paragraphs` - 讀取所有段落（支援過濾）

**Run格式操作 (2)**
- `word_get_run_format` - 讀取Run格式資訊
- `word_set_run_format` - 設定Run格式

**表格操作增強 (3)**
- `word_get_tables` - 讀取所有表格（支援內容）
- `word_copy_table` - 複製表格到其他位置
- `word_move_table` - 移動表格到其他位置

**頁面設定增強 (3)**
- `word_set_page_orientation` - 設定頁面方向（橫向/縱向）
- `word_set_page_margins` - 設定頁邊距
- `word_set_page_size` - 設定頁面大小

**文檔比較與修訂 (4)**
- `word_compare_documents` - 比較兩個文檔並生成比較文檔
- `word_accept_all_revisions` - 接受所有修訂
- `word_reject_all_revisions` - 拒絕所有修訂
- `word_get_revisions` - 讀取所有修訂

**表單欄位 (4)**
- `word_add_form_field` - 添加表單欄位（文字輸入、複選框、下拉選單）
- `word_get_form_fields` - 讀取所有表單欄位
- `word_edit_form_field` - 編輯表單欄位值
- `word_delete_form_field` - 刪除表單欄位

### Excel 表格處理 (121 個)

**基本操作 (5)**
- `excel_create` - 創建工作簿
- `excel_get_content` - 讀取工作簿內容
- `excel_write_cell` - 寫入單元格
- `excel_write_range` - 寫入範圍
- `excel_batch_write` - 批量寫入數據

**工作表操作 (8)**
- `excel_add_sheet` - 添加工作表
- `excel_delete_sheet` - 刪除工作表
- `excel_rename_sheet` - 重命名工作表
- `excel_copy_sheet` - 複製工作表
- `excel_move_sheet` - 移動工作表
- `excel_hide_sheet` - 隱藏/顯示工作表
- `excel_get_sheet_info` - 讀取工作表詳細資訊
- `excel_get_sheets` - 獲取工作表列表

**行列操作 (6)**
- `excel_insert_row` - 插入行
- `excel_delete_row` - 刪除行
- `excel_insert_column` - 插入列
- `excel_delete_column` - 刪除列
- `excel_set_row_height` - 設定行高
- `excel_set_column_width` - 設定列寬

**單元格操作 (2)**
- `excel_merge_cells` - 合併/取消合併單元格
- `excel_get_merged_cells` - 讀取合併單元格資訊

**數據操作 (4)**
- `excel_sort_data` - 排序數據
- `excel_auto_filter` - 自動篩選
- `excel_get_filter_status` - 讀取篩選狀態
- `excel_add_data_validation` - 數據驗證

**格式與圖表 (6)**
- `excel_format_cells` - 格式化單元格
- `excel_copy_format` - 複製單元格格式（格式刷）
- `excel_add_chart` - 添加圖表
- `excel_add_formula` - 添加公式
- `excel_add_pivot_table` - 添加樞紐表
- `excel_add_conditional_formatting` - 添加條件格式

**高級功能 (5)**
- `excel_freeze_panes` - 凍結窗格
- `excel_get_freeze_panes` - 讀取凍結窗格狀態
- `excel_add_hyperlink` - 添加超連結
- `excel_add_image` - 添加圖片
- `excel_set_page_setup` - 頁面設定

**讀取操作 (20)**
- `excel_get_statistics` - 讀取統計資訊
- `excel_get_charts` - 讀取圖表資訊
- `excel_get_pivot_tables` - 讀取樞紐表資訊
- `excel_get_hyperlinks` - 讀取超連結資訊
- `excel_get_named_ranges` - 讀取名稱範圍資訊
- `excel_get_conditional_formatting` - 讀取條件格式資訊
- `excel_get_data_validation` - 讀取數據驗證資訊
- `excel_get_cell_format` - 讀取單元格格式資訊
- `excel_get_formula` - 讀取公式資訊
- `excel_get_images` - 讀取圖片資訊
- `excel_get_protection` - 讀取保護設定資訊
- `excel_get_sheets` - 讀取工作表列表
- `excel_get_merged_cells` - 讀取合併單元格資訊
- `excel_get_cell_value` - 讀取單元格值、公式和類型
- `excel_get_range` - 讀取範圍數據（可選格式資訊）
- `excel_get_workbook_properties` - 讀取工作簿屬性（元數據）
- `excel_get_sheet_properties` - 讀取工作表屬性和設定
- `excel_get_comments` - 讀取批注資訊
- `excel_get_formula_result` - 讀取公式計算結果
- `excel_get_styles` - 讀取樣式資訊（注意：Aspose.Cells不支援命名樣式）

**編輯操作 (9)**
- `excel_edit_chart` - 編輯圖表
- `excel_edit_pivot_table` - 編輯樞紐表
- `excel_edit_conditional_formatting` - 編輯條件格式
- `excel_edit_data_validation` - 編輯數據驗證
- `excel_edit_hyperlink` - 編輯超連結
- `excel_update_chart_data` - 更新圖表數據源
- `excel_edit_cell` - 編輯單元格值和公式
- `excel_edit_range` - 編輯範圍數據
- `excel_edit_sheet_properties` - 編輯工作表屬性（名稱、可見性、標籤顏色等）
- `excel_edit_comment` - 編輯批注

**刪除操作 (10)**
- `excel_delete_chart` - 刪除圖表
- `excel_delete_pivot_table` - 刪除樞紐表
- `excel_delete_hyperlink` - 刪除超連結
- `excel_delete_image` - 刪除圖片
- `excel_delete_conditional_formatting` - 刪除條件格式
- `excel_delete_data_validation` - 刪除數據驗證
- `excel_delete_named_range` - 刪除名稱範圍
- `excel_clear_cell` - 清空單元格內容和/或格式
- `excel_clear_range` - 清空範圍內容和/或格式
- `excel_delete_comment` - 刪除批注

**範圍操作 (2)**
- `excel_copy_range` - 複製範圍（支援複製值、格式、公式）
- `excel_move_range` - 移動範圍到其他位置

**單元格操作增強 (3)**
- `excel_insert_cells` - 插入單元格（向右或向下移動）
- `excel_delete_cells` - 刪除單元格（向左或向上移動）
- `excel_set_cell_locked` - 設定單元格鎖定狀態（用於保護）

**工作表外觀與視圖 (7)**
- `excel_set_sheet_tab_color` - 設定工作表標籤顏色
- `excel_set_gridlines_visible` - 設定網格線顯示/隱藏
- `excel_set_row_column_headers_visible` - 設定行列標題顯示/隱藏
- `excel_set_zero_values_visible` - 設定零值顯示/隱藏
- `excel_set_sheet_background` - 設定工作表背景圖片
- `excel_set_zoom` - 設定工作表縮放比例
- `excel_split_window` - 拆分工作表視窗

**樞紐表欄位操作 (2)**
- `excel_add_pivot_table_field` - 添加樞紐表欄位（行、列、數據、頁面）
- `excel_delete_pivot_table_field` - 刪除樞紐表欄位

**圖表操作增強 (2)**
- `excel_set_chart_title` - 設定圖表標題
- `excel_set_chart_legend` - 設定圖例位置和可見性

**列印設定 (2)**
- `excel_set_print_area` - 設定列印區域
- `excel_set_print_titles` - 設定列印標題（重複行/列）

**公式操作增強 (3)**
- `excel_calculate_all_formulas` - 計算所有公式
- `excel_set_array_formula` - 設定陣列公式
- `excel_get_array_formula` - 讀取陣列公式

**數據驗證訊息 (2)**
- `excel_set_data_validation_error_message` - 設定數據驗證錯誤訊息
- `excel_set_data_validation_input_message` - 設定數據驗證輸入訊息（工具提示）

**實用工具 (2)**
- `excel_get_used_range` - 讀取已使用範圍（數據範圍）
- `excel_get_cell_address` - 單元格地址格式轉換（A1 ↔ 行列索引）

**進階功能 (4)**
- `excel_refresh_pivot_table` - 刷新樞紐表
- `excel_find_replace` - 查找替換
- `excel_calculate_formula` - 計算公式
- `excel_protect_workbook` - 保護工作簿

**文檔操作 (6)**
- `excel_convert` - 轉換格式
- `excel_protect` - 保護工作表
- `excel_unprotect` - 解除工作簿/工作表保護
- `excel_merge_workbooks` - 合併多個工作簿
- `excel_split_workbook` - 拆分工作簿（按工作表）
- `excel_set_workbook_properties` - 設定工作簿屬性（元數據）

**批注操作 (3)**
- `excel_add_comment` - 添加批注
- `excel_get_comments` - 讀取批注
- `excel_delete_comment` - 刪除批注
- `excel_edit_comment` - 編輯批注

**分組操作 (4)**
- `excel_group_rows` - 分組行（建立大綱組）
- `excel_ungroup_rows` - 取消分組行
- `excel_group_columns` - 分組列（建立大綱組）
- `excel_ungroup_columns` - 取消分組列

**樣式操作 (3)**
- `excel_create_style` - 創建樣式（注意：Aspose.Cells不支援命名樣式）
- `excel_apply_style` - 應用樣式到單元格或範圍
- `excel_copy_sheet_format` - 複製工作表格式（列寬、行高等）

**視圖與列印設定 (2)**
- `excel_set_view_settings` - 設定工作表視圖（縮放、網格線、標題等）
- `excel_set_print_settings` - 設定列印設定（列印區域、標題行、方向等）

**批量操作 (1)**
- `excel_batch_format_cells` - 批量格式化多個範圍

### PowerPoint 簡報處理 (97 個)

**基本操作 (6)**
- `ppt_create` - 創建簡報
- `ppt_get_content` - 讀取簡報內容
- `ppt_add_slide` - 添加幻燈片
- `ppt_delete_slide` - 刪除幻燈片
- `ppt_move_slide` - 移動/重排幻燈片
- `ppt_duplicate_slide` - 複製幻燈片
- `ppt_add_text` - 添加文字
- `ppt_add_image` - 添加圖片
- `ppt_add_table` - 添加表格
- `ppt_add_chart` - 添加圖表
- `ppt_add_animation` - 添加動畫
- `ppt_apply_theme` - 應用主題
- `ppt_add_notes` - 添加/更新講者備註
- `ppt_set_background` - 設定幻燈片背景（顏色/圖片）
- `ppt_set_transition` - 設定轉場效果與時間
- `ppt_set_slide_size` - 設定頁面尺寸（自訂或預設）
- `ppt_add_hyperlink` - 插入含超連結文字框
- `ppt_replace_text` - 查找替換文字
- `ppt_extract_images` - 匯出圖片
- `ppt_export_slides_as_images` - 幻燈片轉圖像(整頁)
- `ppt_add_audio` - 插入音訊
- `ppt_add_video` - 插入影片
- `ppt_get_slides_info` - 投影片標題/形狀/備註摘要
- `ppt_set_shape_format` - 設定形狀位置/尺寸/旋轉/填色/線色
- `ppt_set_footer` - 設定頁尾文字、頁碼與日期顯示
- `ppt_set_layout` - 設定投影片版面配置
- `ppt_align_shapes` - 對齊多個形狀（左右上下置中）
- `ppt_reorder_shape` - 調整形狀順序（前/後移）
- `ppt_add_smartart` - 插入 SmartArt 圖形
- `ppt_manage_smartart_nodes` - SmartArt 節點新增/刪除/重命名/移動
- `ppt_batch_set_header_footer` - 批次設定頁尾/頁碼/日期
- `ppt_apply_layout_range` - 批次套用版面配置
- `ppt_copy_shape` - 複製形狀到其他投影片
- `ppt_batch_format_text` - 批次文字樣式設定（字型/大小/粗斜體/顏色）
- `ppt_set_media_playback` - 設定音訊/影片自動或點擊播放、音量、循環
- `ppt_replace_image_with_compression` - 替換圖片並可指定 JPEG 品質
- `ppt_apply_master` - 套用母片及版面到多張投影片
- `ppt_get_sections` - 獲取所有章節與投影片數
- `ppt_hide_slides` - 隱藏/顯示投影片
- `ppt_set_slide_numbering` - 設定起始頁碼
- `ppt_set_shape_hyperlink` - 為任意形狀設定超連結
- `ppt_get_shapes` - 列出形狀資訊（型別/文字/位置/尺寸）
- `ppt_delete_shape` - 刪除指定形狀
- `ppt_get_layouts` - 獲取母片/版面列表
- `ppt_add_section` - 新增章節
- `ppt_rename_section` - 重新命名章節
- `ppt_delete_section` - 刪除章節（可選保留投影片）
- `ppt_clear_notes` - 清空講者備註
- `ppt_convert` - 轉換格式

**編輯操作 (7)**
- `ppt_edit_text` - 編輯文字內容
- `ppt_edit_table` - 編輯表格內容和格式
- `ppt_edit_table_cell` - 編輯表格單元格（內容、格式、字型）
- `ppt_edit_chart` - 編輯圖表（標題、類型）
- `ppt_edit_image` - 編輯圖片（替換、調整大小）
- `ppt_edit_animation` - 編輯動畫效果
- `ppt_edit_hyperlink` - 編輯超連結

**讀取操作 (8)**
- `ppt_get_table_content` - 讀取表格內容
- `ppt_get_chart_data` - 讀取圖表數據和資訊
- `ppt_get_animations` - 讀取動畫資訊
- `ppt_get_hyperlinks` - 讀取所有超連結
- `ppt_get_document_properties` - 讀取文檔屬性（元數據）
- `ppt_get_shape_format` - 讀取形狀格式詳情
- `ppt_get_statistics` - 讀取統計資訊
- `ppt_get_notes` - 讀取講者備註

**讀取操作增強 (7)**
- `ppt_get_slide_details` - 讀取幻燈片詳細資訊（轉場、動畫、背景、備註等）
- `ppt_get_shape_details` - 讀取形狀詳細資訊（位置、大小、旋轉、超連結等）
- `ppt_get_transition` - 讀取轉場資訊
- `ppt_get_master_slides` - 讀取所有母版幻燈片及其版式
- `ppt_get_layouts` - 讀取所有版式資訊
- `ppt_get_background` - 讀取幻燈片背景資訊
- `ppt_get_protection` - 讀取保護資訊

**編輯操作增強 (6)**
- `ppt_edit_slide` - 編輯幻燈片屬性（隱藏狀態、備註等）
- `ppt_edit_shape` - 編輯形狀屬性（位置、大小、旋轉、翻轉等）
- `ppt_edit_notes` - 編輯講者備註
- `ppt_set_document_properties` - 設定文檔屬性（元數據）
- `ppt_set_header` - 設定頁眉文字
- `ppt_set_slide_orientation` - 設定幻燈片方向（橫向/縱向）

**刪除操作增強 (2)**
- `ppt_clear_slide` - 清空幻燈片所有形狀
- `ppt_delete_transition` - 刪除轉場效果

**刪除操作 (6)**
- `ppt_delete_table` - 刪除表格
- `ppt_delete_chart` - 刪除圖表
- `ppt_delete_animation` - 刪除動畫
- `ppt_delete_hyperlink` - 刪除超連結
- `ppt_delete_audio` - 刪除音訊
- `ppt_delete_video` - 刪除影片

**文檔操作 (5)**
- `ppt_merge` - 合併多個演示文稿
- `ppt_split` - 拆分演示文稿
- `ppt_protect` - 保護演示文稿（密碼）
- `ppt_unprotect` - 解除保護
- `ppt_set_properties` - 設定文檔屬性

**形狀操作 (3)**
- `ppt_group_shapes` - 組合多個形狀
- `ppt_ungroup_shapes` - 取消組合形狀
- `ppt_flip_shape` - 翻轉形狀（水平/垂直）

**表格操作 (4)**
- `ppt_insert_table_row` - 插入表格行（API限制，可能需要重建表格）
- `ppt_insert_table_column` - 插入表格列（API限制，可能需要重建表格）
- `ppt_delete_table_row` - 刪除表格行
- `ppt_delete_table_column` - 刪除表格列

**圖表操作 (1)**
- `ppt_update_chart_data` - 更新圖表數據（結構準備，完整實現可能需要圖表特定邏輯）

### PDF 文件處理 (47 個)

**基本操作 (2)**
- `pdf_create` - 創建PDF
- `pdf_get_content` - 獲取PDF內容

**內容添加 (5)**
- `pdf_add_text` - 添加文字
- `pdf_add_image` - 添加圖片
- `pdf_add_table` - 添加表格
- `pdf_add_watermark` - 添加浮水印
- `pdf_add_page` - 添加頁面

**書籤與註解 (2)**
- `pdf_add_bookmark` - 添加書籤
- `pdf_add_annotation` - 添加註解

**編輯操作 (4)**
- `pdf_edit_text` - 編輯文字（替換）
- `pdf_edit_table` - 編輯表格單元格
- `pdf_edit_bookmark` - 編輯書籤屬性
- `pdf_edit_annotation` - 編輯註解屬性

**讀取操作 (6)**
- `pdf_extract_text` - 提取文字
- `pdf_extract_images` - 提取圖片
- `pdf_get_page_info` - 讀取頁面資訊
- `pdf_get_bookmarks` - 讀取書籤列表
- `pdf_get_annotations` - 讀取註解列表
- `pdf_get_document_properties` - 讀取文檔屬性（元數據）
- `pdf_get_form_fields` - 讀取表單欄位
- `pdf_get_statistics` - 讀取統計資訊

**刪除操作 (8)**
- `pdf_delete_page` - 刪除頁面
- `pdf_delete_bookmark` - 刪除書籤
- `pdf_delete_annotation` - 刪除註解
- `pdf_delete_form_field` - 刪除表單欄位
- `pdf_delete_link` - 刪除超連結
- `pdf_delete_attachment` - 刪除附件
- `pdf_delete_signature` - 刪除數字簽名
- `pdf_delete_image` - 刪除圖片（從資源中移除）

**編輯操作增強 (3)**
- `pdf_edit_form_field` - 編輯表單欄位屬性（值、位置、大小等）
- `pdf_edit_link` - 編輯超連結屬性（URL、目標頁面、位置、大小等）
- `pdf_edit_image` - 編輯圖片屬性（位置、大小、旋轉等，需要內容流操作）

**讀取操作增強 (4)**
- `pdf_get_links` - 讀取所有超連結
- `pdf_get_attachments` - 讀取所有附件
- `pdf_get_signatures` - 讀取所有數字簽名
- `pdf_get_page_details` - 讀取頁面詳細資訊（大小、旋轉、註解、圖片等）

**附件操作 (1)**
- `pdf_add_attachment` - 添加附件（文件）

**進階功能 (1)**
- `pdf_redact` - 編輯（塗黑）文字或區域

**頁面操作 (1)**
- `pdf_rotate_page` - 旋轉頁面

**連結與表單 (2)**
- `pdf_add_link` - 添加超連結
- `pdf_add_form_field` - 添加表單欄位（文字框、複選框、單選按鈕）

**文檔操作 (4)**
- `pdf_merge` - 合併PDF
- `pdf_split` - 拆分PDF
- `pdf_encrypt` - 加密PDF
- `pdf_sign` - 簽章PDF
- `pdf_set_document_properties` - 設定文檔屬性
- `pdf_compress` - 壓縮PDF

### 轉換工具

轉換工具已**自動集成**到各個文檔工具中，無需額外配置：

**各工具專屬轉換**：
- `word_convert` - Word格式轉換（啟用 `--word` 時可用）
- `excel_convert` - Excel格式轉換（啟用 `--excel` 時可用）
- `ppt_convert` - PowerPoint格式轉換（啟用 `--ppt` 時可用）

**通用轉換工具**：
- `convert_to_pdf` - 將任何文檔轉換為PDF（啟用任何文檔工具時自動可用）
- `convert_document` - 跨格式轉換（啟用兩個或以上文檔工具時自動可用）

**使用範例**：
- 只啟用 `--excel`：可使用 `excel_convert` 和 `convert_to_pdf`
- 啟用 `--word --excel`：可使用所有Word/Excel轉換工具 + `convert_to_pdf` + `convert_document`（支援Word↔Excel互轉）
- 啟用 `--all`：可使用所有轉換功能

## 🎉 主要特性

### MCP 2025-11-25 規範支持
- ✅ 符合最新 MCP 協議規範（protocolVersion: 2025-11-25）
- ✅ 自動工具注解（readonly/destructive）基於命名約定
- ✅ 完整的 JSON-RPC 2.0 錯誤處理

### 統一字型設定
多個工具支援中英文字型分別設定（`fontNameAscii` 和 `fontNameFarEast` 參數）

### 靈活的授權配置
- 支援總授權或單一組件授權
- 自動搜尋、環境變數或命令列參數配置
- 試用模式降級（找不到授權時）

## 📝 使用範例

### 從A文檔複製格式到B文檔

**複製段落格式：**
```
1. word_get_paragraph_format(path="A.docx", paragraphIndex=0)
2. 使用返回的格式信息
3. word_edit_paragraph(path="B.docx", paragraphIndex=0, ...)
```

**複製表格結構：**
```
1. word_get_table_structure(path="A.docx", tableIndex=0)
2. 參考返回的結構信息
3. word_add_table(path="B.docx", ...) 創建相同結構
```

**複製樣式：**
```
word_copy_styles_from(sourcePath="A.docx", targetPath="B.docx")
```

## 🌍 跨平台支持

所有平台由 **GitHub Actions** 自動構建和發布：
- ✅ Windows (x64)
- ✅ Linux (x64)
- ✅ macOS Intel (x64)
- ✅ macOS ARM (arm64 - M1/M2/M3)

**獲取方式：** 從 [GitHub Releases](../../releases) 下載最新版本

## 📄 授權

本項目需要有效的 Aspose 授權文件。支援以下授權類型：
- `Aspose.Total.lic` - 總授權（包含所有組件）
- `Aspose.Words.lic`、`Aspose.Cells.lic`、`Aspose.Slides.lic`、`Aspose.Pdf.lic` - 單一組件授權

**配置方式：**
1. 將授權文件放在可執行文件同一目錄（自動搜尋）
2. 使用環境變數 `ASPOSE_LICENSE_PATH` 指定路徑
3. 使用命令列參數 `--license:路徑` 指定路徑

如果找不到授權文件，系統會以試用模式運行（會有試用版標記）。

## 🔗 相關資源

- [Aspose.Total for .NET](https://products.aspose.com/total/net/)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [Claude Desktop](https://claude.ai/desktop)
