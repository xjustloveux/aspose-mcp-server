# Aspose MCP Server — Codex 自訂市集

本外掛讓 Codex 透過 stdio 使用本機 Aspose MCP Server，讀取、編輯及轉換文件。
外掛套件提供 MCP 設定；執行檔需另外下載，不需要安裝 .NET Runtime。

## 1. 準備本機執行檔

從 [GitHub Releases](https://github.com/xjustloveux/aspose-mcp-server/releases)
下載對應平台的執行檔封存：

| 平台 | 封存檔 |
|---|---|
| Windows x64 | `aspose-mcp-server-windows-x64.zip` |
| Linux x64 | `aspose-mcp-server-linux-x64.tar.gz` |
| macOS Intel | `aspose-mcp-server-macos-x64.tar.gz` |
| macOS ARM64 | `aspose-mcp-server-macos-arm64.tar.gz` |

解壓縮到固定位置，將含 `AsposeMcpServer.exe`（Windows）或 `AsposeMcpServer`
（Linux/macOS）的資料夾加入作業系統使用者的 `PATH`。Windows 可在「環境變數」
編輯使用者的 Path；macOS/Linux 需確保啟動 Codex 的環境也有此 PATH，只有互動式
shell 的設定可能不會傳給桌面程式。完整關閉並重開 Codex，讓它取得更新後的環境。

在同一啟動環境確認可找到執行檔：

```powershell
# Windows
Get-Command AsposeMcpServer
```

```bash
# macOS / Linux
command -v AsposeMcpServer
```

macOS 的執行權限及隔離標記處理請參閱專案
[部署指南](https://xjustloveux.github.io/aspose-mcp-server/deployment.html)。

## 2. 加入市集並安裝

發行提供兩種 ZIP，兩者都不包含執行檔或使用者的 Aspose 授權：

| 封存檔 | 用途 |
|---|---|
| `aspose-mcp-server-codex-marketplace.zip` | 解壓縮後新增市集，再安裝其中的外掛 |
| `aspose-mcp-server-codex-plugin.zip` | 單一外掛封存，供「上傳外掛封存檔」入口選取 |

單一外掛 ZIP 的根目錄包含 `plugin.json`、`mcp.json`、`.codex-plugin/`、
`README.md` 與 MIT `LICENSE`，不包含市集清單。上傳入口的本機 stdio MCP
相容性尚未實測；官方提交入口的 Skills only 流程會排除 MCP 設定，不能用來
安裝此 MCP 外掛。若介面拒絕本機 MCP 套件，請使用下方已驗證的自訂市集方式。

使用專案原始碼時，在專案根目錄執行以下指令。使用發行的
`aspose-mcp-server-codex-marketplace.zip` 時，先解壓縮，再進入包含
`.agents/` 與 `plugins/` 的解壓縮根目錄執行：

```text
codex plugin marketplace add .
```

在 Codex 的外掛目錄選擇 **Aspose Local Document Tools**，安裝 **Aspose MCP Server**，
然後開啟新聊天使用。安裝只提供外掛設定；執行檔與授權仍由本機環境提供。

若市集未顯示，完整重啟 Codex，並在上述根目錄檢查：

```text
codex plugin marketplace list
codex plugin list --marketplace aspose-local --available --json
```

專案中的市集檔為 `.agents/plugins/marketplace.json`，其
`source.path` 相對於市集根目錄解析。套件提供根目錄 `plugin.json` 與 `mcp.json`，
並附 `.codex-plugin/plugin.json` 相容 manifest，供 Codex 0.146 等版本讀取。
兩種 manifest 使用同一份 MCP 設定，發行打包時會同步設定版本號。

## 3. 授權、工具類別與文件

在作業系統使用者環境設定 `ASPOSE_LICENSE_PATH` 為自己的 Aspose 授權檔絕對路徑，
再重開 Codex。也可以將授權檔放在執行檔旁，沿用伺服器的自動搜尋功能。
沒有授權時使用評估模式，輸出可能有浮水印及其他 Aspose 評估限制。
專案原始碼的 MIT 授權與 Aspose 元件授權分別適用；市集 ZIP 不包含 Aspose 授權檔。

預設啟用所有文件類別。可透過 `ASPOSE_TOOLS=word,excel,pdf` 等使用者環境設定
選擇所需類別；外掛只傳入 `--stdio`，並透過 Codex 的 `env_vars` 明確轉送
`ASPOSE_LICENSE_PATH` 與 `ASPOSE_TOOLS`。這些變數必須存在於啟動 Codex 的環境中。
Session 與 Extension 沿用伺服器預設設定。

傳給工具的文件路徑必須是 MCP 執行環境可存取的絕對路徑。遠端主機、WSL 與本機
Windows 的檔案系統各自不同；外掛安裝不會自動搬移文件或改變主機檔案權限。

## 開發與更新

在專案根目錄執行（Python 3.10+，只使用標準函式庫）：

```text
python -m unittest discover -s deploy -p test_pack_codex.py -v
python deploy/pack-codex.py --version 0.1.0 --output publish/aspose-mcp-server-codex-marketplace.zip
python deploy/pack-codex.py --format plugin --version 0.1.0 --output publish/aspose-mcp-server-codex-plugin.zip
```

輸出路徑必須尚不存在。打包只讀取明確列出的市集設定、manifest、MCP 設定、
本說明與 MIT LICENSE；不會修改原始碼中的版本或收集工作目錄的其他檔案。
CI 使用同次發行的版本號產生 ZIP，與平台執行檔一起發布。

修改外掛來源後，刷新市集並重新安裝／重新啟用外掛，再開啟新聊天確認工具載入。
Codex 使用安裝快取中的外掛副本，不直接執行這個來源目錄的設定。

封裝格式與市集操作依據 [OpenAI 官方文件](https://developers.openai.com/plugins/build/plugins)。
