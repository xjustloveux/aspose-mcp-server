# Code Quality Exceptions Documentation

本文件記錄 JetBrains InspectCode 報告中被排除修復的問題及其原因。
這些問題經過評估後決定保留，未來進行代碼品質檢查時可參考本文件跳過這些項目。

**最後更新日期**: 2026-09-10
**分析工具**: JetBrains InspectCode 2025.3.3、SonarCloud、.NET analyzers

> 本文件不記錄行號。行號在重構後會失效而不會被任何檢查發現，
> 因此例外一律以「檔案 + 符號」定位。
> `Tests/Infrastructure/CodeQualityExceptionsDocTests` 會驗證本文件引用的每一個
> 檔案都存在，並禁止重新加入行號欄位。

---

## 目錄

1. [AccessToDisposedClosure](#1-accesstodisposedclosure)
2. [AutoPropertyCanBeMadeGetOnly.Global](#2-autopropertycanbemadegetonlyglobal)
3. [ClassNeverInstantiated.Global](#3-classneverinstantiatedglobal)
4. [CompareOfFloatsByEqualityOperator](#4-compareoffloatsbyequalityoperator)
5. [ConvertToPrimaryConstructor](#5-converttoprimaryconstructor)
6. [MemberCanBePrivate.Global](#6-membercanbeprivateglobal)
7. [MemberCanBeProtected.Global](#7-membercanbeprotectedglobal)
8. [MethodSupportsCancellation](#8-methodsupportscancellation)
9. [ParameterOnlyUsedForPreconditionCheck.Local](#9-parameteronlyusedforpreconditionchecklocal)
10. [PropertyCanBeMadeInitOnly.Global](#10-propertycanbemadeinitonlyglobal)
11. [UnusedMember.Global](#11-unusedmemberglobal)
12. [UnusedMethodReturnValue.Global](#12-unusedmethodreturnvalueglobal)
13. [UnusedType.Global](#13-unusedtypeglobal)
14. [UseObjectOrCollectionInitializer](#14-useobjectorollectioninitializer)
15. [UseUtf8StringLiteral](#15-useutf8stringliteral)
16. [MethodHasAsyncOverload](#16-methodhasasyncoverload)
17. [ClassNeverInstantiated.Local](#17-classneverinstantiatedlocal)
18. [SonarCloud 與 .NET analyzer 精確例外](#18-sonarcloud-與-net-analyzer-精確例外)

---

## 1. AccessToDisposedClosure

| 項目 | 內容 |
|------|------|
| **級別** | Warning |
| **範圍** | 測試專案，以及 8 個同步 writer callback |
| **訊息** | Captured variable is disposed in the outer scope |
| **處理方式** | 測試由 `.editorconfig` 排除；production 僅限精確檔案或單行排除 |

### 受影響檔案

| 檔案 | 變量 |
|------|------|
| `Tests/Core/Session/DocumentContextTests.cs` | `context` |
| `Tests/Core/Session/DocumentSessionManagerTests.cs` | `manager` |
| `Tests/Core/Session/DocumentSessionManagerTests.cs` | `manager` |
| `Tests/Core/Session/DocumentSessionManagerTests.cs` | `manager` |
| `Core/Conversion/DocumentConverter.cs` | 7 個 `pdfDoc` writer callback |
| `Handlers/PowerPoint/Image/ExportSlidesHandler.cs` | `bmp` writer callback |

### 問題描述

在 lambda 或匿名方法中捕獲的變量在外部作用域被 dispose。

### 不修復原因

- 這是測試代碼中的模式，用於測試 `Assert.Throws` 等異常處理場景
- 這些是故意設計的測試場景，用於驗證對象 dispose 後的行為
- 修改可能導致測試無法正確驗證預期行為
- 測試需要驗證「已釋放對象被存取時應拋出異常」的行為
- `BoundedFilePublisher.Publish` 與 `BoundedFileBatch.Stage` 都會在方法返回前同步執行 writer；
  delegate 不會被保存或逸出，因此 production 的 8 個位置不會在資源釋放後才執行
- WebSocket 原有的 callback 確實可能在 deadline 被釋放後執行，並未套用例外；該處已改為
  `RefreshableCancellationDeadline`，以同步方式協調 refresh 與 dispose

### 範例代碼

```csharp
// 測試代碼故意在 dispose 後存取對象以驗證異常行為
using var document = new Document();
// ... 設置測試
document.Dispose();
Assert.Throws<ObjectDisposedException>(() => document.SomeMethod());
```

---

## 2. AutoPropertyCanBeMadeGetOnly.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 3 |
| **訊息** | Auto-property can be made get-only |

### 受影響檔案

| 檔案 | 屬性 |
|------|------|
| `Core/Security/JwtConfig.cs` | `ClientSecret` (JWT 配置屬性) |
| `Core/Security/JwtConfig.cs` | `CustomEndpoint` (JWT 配置屬性) |
| `Core/Security/ApiKeyConfig.cs` | `CustomEndpoint` (API Key 配置屬性) |

### 問題描述

自動屬性可以改為唯讀（移除 setter）。

### 不修復原因

- 這些是配置類屬性，需要支援 JSON 反序列化
- `System.Text.Json` 和其他序列化器需要 `set` 存取器來設置值
- 移除 `set` 會導致從 `config.json` 載入配置失敗

### 範例代碼

```csharp
// 配置類需要 setter 以支援 JSON 反序列化
public class AuthConfig
{
    public string ApiKey { get; set; }  // 需要 set 來載入 JSON
}

// 使用範例
var config = JsonSerializer.Deserialize<AuthConfig>(json);
```

---

## 3. ClassNeverInstantiated.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Class is never instantiated |
| **處理方式** | 已添加 `// ReSharper disable once` 註解 |

### 受影響檔案

| 檔案 | 類別 |
|------|------|
| `Tests/Core/Handlers/HandlerRegistryAutoDiscoveryTests.cs` | `DifferentContextDocument` |

### 問題描述

類別定義了但從未被 `new` 實例化。

### 不修復原因

- 這是測試用的類別，用於測試「當 Handler 的泛型參數是不同 Context 類型時，不會被自動發現」
- 類別只需要存在作為泛型類型參數，不需要實際實例化
- 這是有意為之的測試設計

### 範例代碼

```csharp
// 這個類別用於驗證 DifferentContextHandler 不會被發現
public class DifferentContextDocument
{
    public int Value { get; set; }
}

// 使用此類別的 Handler 不應被 HandlerRegistry<TestDiscoveryDocument> 發現
public class DifferentContextHandler : OperationHandlerBase<DifferentContextDocument>
{
    // ...
}
```

---

## 4. CompareOfFloatsByEqualityOperator

| 項目 | 內容 |
|------|------|
| **級別** | Warning |
| **數量** | 1 |
| **訊息** | Comparison of floating point numbers with equality operator |
| **處理方式** | 已添加 `// ReSharper disable once` 註解 |

### 受影響檔案

| 檔案 | 說明 |
|------|------|
| `Tests/Handlers/Excel/DataOperations/GetContentHandlerTests.cs` | 整數值 100 的浮點比較 |

### 問題描述

使用 `==` 運算符比較浮點數可能因精度問題導致意外結果。

### 不修復原因

這是測試代碼中驗證 Excel 儲存格值的邏輯。比較的值是整數 `100`，不是浮點運算結果。
當整數被存儲為 `double` 類型時（Excel 的內部表示），精確的整數值比較是安全的。

### 範例代碼

```csharp
// 測試代碼 - 檢查儲存格值是否為 100
// ReSharper disable once CompareOfFloatsByEqualityOperator - Exact integer value 100 comparison is safe
Assert.Contains(result.Rows[1].Values,
    v => v?.ToString() == "100" || v is (int or double) and 100);
```

整數 100 可以精確表示為 IEEE 754 double，因此這個比較是安全的。

---

## 5. ConvertToPrimaryConstructor

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Convert into primary constructor |
| **處理方式** | 已在 `.editorconfig` 中全局排除 |

### 受影響檔案

| 檔案 | 類別 |
|------|------|
| `Tests/Core/Handlers/HandlerRegistryAutoDiscoveryTests.cs` | `NoParameterlessCtorHandler` |

### 問題描述

建議將傳統建構函式轉換為 C# 12 的主要建構函式語法。

### 不修復原因

- 這是 C# 12 引入的語法糖，純屬風格選擇
- 團隊決定保持傳統建構函式寫法以維持一致性
- 測試類別 `NoParameterlessCtorHandler` 故意使用帶參數的建構函式來測試「沒有無參數建構函式的類別不會被自動發現」

### 範例代碼

```csharp
// 目前寫法（保留）
public class NoParameterlessCtorHandler : OperationHandlerBase<TestDiscoveryDocument>
{
    private readonly string _requiredValue;

    public NoParameterlessCtorHandler(string requiredValue)
    {
        _requiredValue = requiredValue;
    }
}

// C# 12 主要建構函式寫法（不採用）
public class NoParameterlessCtorHandler(string requiredValue) : OperationHandlerBase<TestDiscoveryDocument>
{
    private readonly string _requiredValue = requiredValue;
}
```

---

## 6. MemberCanBePrivate.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 6 |
| **訊息** | Member can be made private |

### 受影響檔案

| 檔案 | 成員 |
|------|------|
| `Core/Transport/TransportConfig.cs` | `Mode.set` |
| `Core/Transport/TransportConfig.cs` | `Port.set` |
| `Core/Transport/TransportConfig.cs` | `Host.set` |
| `Core/Session/DocumentSession.cs` | `LastAccessedAt.set` |
| `Core/Tracking/TrackingConfig.cs` | `WebhookAuthHeader.set` |
| `Core/Tracking/TrackingConfig.cs` | `WebhookTimeoutSeconds.set` |

### 問題描述

成員可以改為 private 可見性。

### 不修復原因

- 這些是公開 API 的一部分，外部可能需要存取
- 配置類屬性需要公開 setter 支援 JSON 反序列化
- 降低可見性可能破壞向後相容性

---

## 7. MemberCanBeProtected.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Member can be made protected |

### 受影響檔案

| 檔案 | 成員 |
|------|------|
| `Tests/Infrastructure/TestBase.cs` | `AsposeLibraryType` enum |

### 問題描述

Enum 可以改為 protected 可見性。

### 不修復原因

- 這是測試基類中的公開 enum，供所有測試類使用
- 改為 protected 會限制其在非衍生類中的使用
- 測試程式碼中保持 public 更具彈性

---

## 8. MethodSupportsCancellation

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Method has overload with cancellation support |
| **處理方式** | 測試檔案已在 `.editorconfig` 中排除；生產代碼已添加 `// ReSharper disable once` 註解 |

### 受影響檔案

| 檔案 | 方法 | 備註 |
|------|------|------|
| `Core/Transport/WebSocketConnectionHandler.cs` | `WaitAsync` | 生產代碼 |

### 問題描述

方法有支援 CancellationToken 的重載可用。

### 不修復原因

**生產代碼（WebSocketConnectionHandler）**：
- `WaitAsync(TimeSpan)` 有 `WaitAsync(TimeSpan, CancellationToken)` 重載
- 此處用於等待任務完成的超時控制，已有 `linkedCts` 處理取消
- 添加額外的 CancellationToken 參數會使代碼更複雜
- 目前的實現已經可以正確處理取消場景

---

---

## 9. ParameterOnlyUsedForPreconditionCheck.Local

| 項目 | 內容 |
|------|------|
| **級別** | Warning |
| **數量** | 2 |
| **訊息** | Parameter is only used for precondition check(s) |
| **處理方式** | 已添加 `// ReSharper disable once` 註解 |

### 受影響檔案

| 檔案 | 參數 |
|------|------|
| `Tests/Helpers/PowerPoint/PptLayoutHelperTests.cs` | `item` (Assert.All lambda 參數) |
| `Tests/Handlers/Word/Text/SearchWordTextHandlerTests.cs` | `m` (Assert.All lambda 參數) |

### 問題描述

Lambda 參數僅用於前置條件檢查（如 `Assert` 語句），而非其他邏輯。

### 不修復原因

這是 xUnit 的 `Assert.All` 方法的正確使用方式。該方法需要一個 lambda 來對集合中的每個元素執行驗證。
參數在 lambda 內被用於 Assert 語句，這正是預期的行為。

### 範例代碼

```csharp
// Assert.All 的正確用法 - 參數用於驗證每個元素
// ReSharper disable once ParameterOnlyUsedForPreconditionCheck.Local - Assert.All parameter is intended for validation
Assert.All(result, item =>
{
    Assert.NotNull(item);
    Assert.IsType<GetLayoutInfo>(item);
    Assert.NotNull(item.Type);
});
```

---

## 10. PropertyCanBeMadeInitOnly.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 38 |
| **訊息** | Property can be made init-only |

### 受影響檔案

| 檔案 | 涉及屬性數量 |
|------|-------------|
| `Core/Security/ApiKeyAuthenticationMiddleware.cs` | 多個配置屬性 |
| `Core/Security/JwtAuthenticationMiddleware.cs` | 多個配置屬性 |
| `Core/Security/AuthConfig.cs` | 多個配置屬性 |
| `Core/Tracking/TrackingConfig.cs` | 多個配置屬性 |
| `Core/Session/SessionConfig.cs` | 多個配置屬性 |
| `Core/Session/DocumentSession.cs` | 多個狀態屬性 |
| `Core/Session/DocumentSessionManager.cs` | 多個狀態屬性 |
| `Core/Session/TempFileManager.cs` | 多個配置屬性 |
| `Core/Transport/TransportConfig.cs` | 多個配置屬性 |
| `Core/ServerConfig.cs` | 多個配置屬性 |

### 問題描述

屬性可以改為 init-only（`init` 而非 `set`）。

### 不修復原因

- 這些配置類需要支援 JSON 反序列化
- `System.Text.Json` 預設不支援 `init` 屬性（需要特殊配置）
- 專案可能需要在運行時修改配置值
- 改為 `init` 會破壞現有的配置載入邏輯

### 範例代碼

```csharp
// 使用 set 以支援標準 JSON 反序列化
public string Host { get; set; } = "localhost";

// 如果改為 init，需要額外配置才能反序列化
public string Host { get; init; } = "localhost";  // 會導致反序列化失敗
```

---

## 11. UnusedMember.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 7 |
| **訊息** | Member is never used |

### 受影響檔案

| 檔案 | 成員 |
|------|------|
| `Tests/Infrastructure/ExcelTestBase.cs` | `AssertCellValue()` |
| `Tests/Infrastructure/WordTestBase.cs` | `AssertParagraphExists()` |
| `Tests/Infrastructure/WordTestBase.cs` | `AssertParagraphStyle()` |
| `Tests/Infrastructure/PdfTestBase.cs` | `IsEvaluationMode()` |
| `Core/Session/DocumentSession.cs` | `GetDocumentAsync()` |
| `Core/Session/DocumentSessionManager.cs` | `OnServerShutdown()` |
| `Core/Tracking/TrackingExtensions.cs` | `UseTracking()` |

### 問題描述

成員在專案內部未使用。

### 不修復原因

- 這些是公開 API 的一部分，可能被外部使用者調用
- 測試輔助方法保留供未來測試使用
- 擴展方法 (`UseXxx`) 是設計給外部使用的 API
- 刪除會破壞 API 相容性

---

## 12. UnusedMethodReturnValue.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Return value of method is never used |
| **處理方式** | 已添加 `// ReSharper disable once` 註解 |

### 受影響檔案

| 檔案 | 方法 |
|------|------|
| `Core/McpServerBuilderExtensions.cs` | `WithFilteredTools()` |

### 問題描述

方法返回值從未被使用。

### 不修復原因

- 公開 API 方法，外部可能使用返回值進行鏈式調用
- 移除返回值會破壞 API 簽名
- 符合 Builder Pattern 的設計慣例

### 範例代碼

```csharp
// Builder Pattern 允許鏈式調用
builder
    .WithFilteredTools(filter)
    .WithOtherOption();

// 或獨立使用（目前專案內的用法）
builder.WithFilteredTools(filter);
```

---

## 13. UnusedType.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Type is never used |

### 受影響檔案

| 檔案 | 類型 |
|------|------|
| `Core/Tracking/TrackingExtensions.cs` | `TrackingExtensions` |

### 問題描述

類型在專案內部未使用。

### 不修復原因

- 公開 API 類型，供外部使用者使用
- 這是 ASP.NET Core 擴展方法類，用於 `IApplicationBuilder` 配置
- 專案內部使用不同的配置方式，但保留給外部使用者
- 刪除會破壞外部依賴

### 範例代碼

```csharp
// 外部使用者可以這樣配置
app.UseTracking();
```

---

## 14. UseObjectOrCollectionInitializer

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 26 |
| **訊息** | Use object initializer |

### 受影響檔案

#### 測試檔案 - AuthConfigTests.cs (14 個)

這些是測試程式碼中對 `AuthConfig` 嵌套屬性的設定：

| 檔案 |
|------|
| `Tests/Core/Security/AuthConfigTests.cs` |

#### 測試檔案 - GetParagraphFormatWordHandlerTests.cs (5 個)

| 檔案 | 說明 |
|------|------|
| `Tests/Handlers/Word/Paragraph/GetParagraphFormatWordHandlerTests.cs` | 誤判 |

#### 測試檔案 - AddTableOfContentsWordHandlerTests.cs (1 個)

| 檔案 | 說明 |
|------|------|
| `Tests/Handlers/Word/Reference/AddTableOfContentsWordHandlerTests.cs` | 誤判 |

### 問題描述

建議使用物件初始化器語法。

### 不修復原因

#### AuthConfigTests.cs (測試程式碼可讀性)

測試程式碼中逐步設定屬性更清晰易讀，便於除錯：

```csharp
// 目前寫法 - 清晰的逐步設定
var config = new AuthConfig();
config.ApiKey.Enabled = true;
config.ApiKey.Mode = ApiKeyMode.Local;
config.ApiKey.Keys = ["key1", "key2"];

// 建議的寫法 - 對測試來說較不直觀
var config = new AuthConfig
{
    ApiKey = { Enabled = true, Mode = ApiKeyMode.Local, Keys = ["key1", "key2"] }
};
```

#### GetParagraphFormatWordHandlerTests.cs & AddTableOfContentsWordHandlerTests.cs (誤判)

**這是工具的誤判**。這些代碼是在修改已存在物件的屬性，不是在初始化新物件。

```csharp
// 目前代碼 - 修改 builder.Font 物件的屬性
var builder = new DocumentBuilder(doc);
builder.Font.Bold = true;       // Font 是 builder 的屬性，不是新物件
builder.Font.Italic = true;
builder.Font.Size = 14;

// 工具錯誤建議（這樣寫是不正確的）
var builder = new DocumentBuilder(doc)
{
    Font = { Bold = true }  // 錯誤：Font 是唯讀屬性，不能用初始化器
};
```

`DocumentBuilder.Font` 是一個已存在的物件屬性，我們在修改它的子屬性，這與物件初始化器的使用場景不同。

---

## 15. UseUtf8StringLiteral

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 2 |
| **訊息** | Collection expression can be converted to a UTF-8 string literal |
| **處理方式** | 已添加 `// ReSharper disable UseUtf8StringLiteral` / `// ReSharper restore UseUtf8StringLiteral` 註解對 (行 103-113) |

### 受影響檔案

| 檔案 | 說明 |
|------|------|
| `Tests/Core/ShapeDetailProviders/PictureFrameDetailProviderTests.cs` | 誤判 (PNG 二進制數據) |

### 問題描述

建議將 byte 陣列轉換為 UTF-8 字串字面量。

### 不修復原因

**這是工具的誤判**。這些 byte 陣列是 PNG 圖片格式的二進制數據（magic bytes），不是 UTF-8 編碼的文字。

```csharp
// 目前代碼 - PNG 檔案的二進制標頭
ms.Write([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);  // PNG signature
ms.Write([0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52]);  // IHDR chunk

// 工具錯誤建議（這樣轉換是不正確的）
ms.Write("\x89PNG\r\n\x1a\n"u8);  // 0x89 不是有效的 UTF-8 字元
```

PNG 檔案的 signature 包含 `0x89` 等非 ASCII 字元，這些是二進制數據，無法正確轉換為 UTF-8 字串字面量。

---

## 16. MethodHasAsyncOverload

| 項目 | 內容 |
|------|------|
| **級別** | Suggestion |
| **範圍** | `Tests/**/*.cs` |
| **訊息** | Method has an async overload |
| **處理方式** | 由 `.editorconfig` 排除 |

### 不修復原因

InspectCode 會在同步測試配置與 fixture helper 呼叫具有 async overload 的 API 時提出建議。
這些呼叫刻意保持同步，以確保配置、資源建立與 assertion 的順序明確；真正的非同步
production contract 仍由 async 測試覆蓋。將每個 fixture 機械式改成 async 只會增加狀態機與
測試噪音，不會改善受測行為。

---

## 17. ClassNeverInstantiated.Local

| 項目 | 內容 |
|------|------|
| **級別** | Suggestion |
| **數量** | 1 |
| **訊息** | Record is never instantiated |
| **處理方式** | 精確單行 `ReSharper disable once` |

### 受影響檔案

| 檔案 | 類別 |
|------|------|
| `Tests/Infrastructure/PublishedDocsInventoryTests.cs` | `PendingEntry` |

### 不修復原因

`PendingEntry` 是 `System.Text.Json` 反序列化 manifest 時以反射建立的 DTO。測試程式碼不應為了
讓靜態分析器看見 `new PendingEntry(...)` 而加入不會參與 assertion 的假實例。

---

## 18. SonarCloud 與 .NET analyzer 精確例外

這些項目不是以「清掉數字」為目的跳過，而是經過行為與安全邊界確認後，保留既有契約或
必要的不變量。可縮小到符號的例外均使用 `SuppressMessage`；只有橫跨整個專案、且每個位置
理由一致的 analyzer 建議才放在專案檔的 `NoWarn`。

| 規則 | 檔案 / 符號 | 保留原因 | 處理方式 |
|------|-------------|----------|----------|
| MCP002 | `AsposeMcpServer.csproj`、所有 MCP tool 類別 | 118 個 tool 已有人工維護的公開 `[Description]` 契約；改成 source-generated XML description 會重寫對外工具文案，而不是改善執行行為 | production 專案 `NoWarn` |
| SYSLIB1045 | `Tests/AsposeMcpServer.Tests.csproj`、測試用 Regex | 13 個短小、具 timeout 的測試解析器不在 production runtime；為它們拆成 partial 測試類別不會改善產品效能 | test 專案 `NoWarn` |
| S107 | `Core/Session/DocumentContext.cs` 的私有建構式 | 這是 file/session context 的單一組合點；參數分別表達所有權、身分、session 與 allowlist，包成參數物件只會把必要不變量藏起來 | 符號級 `SuppressMessage` |
| S3264 | `Core/Session/DocumentSessionManager.cs` 的 `SessionClosed` | 事件不是透過 `?.Invoke` 一次呼叫，而是由 `NotifySessionClosed` 逐一執行 invocation list，避免一個失敗的 subscriber 阻斷其他 subscriber | 符號級 `SuppressMessage` |
| S2696 / S3877 | `Helpers/SlidesGate.cs` 的 scope `Dispose` | scope 必須更新 process-wide gate 的 thread-static 深度；跨執行緒 dispose 會破壞計數，因此必須在釋放錯誤執行緒的 hold 前明確失敗 | 符號級 `SuppressMessage` |
| S5443 | `Core/Conversion/DocumentConverter.cs` 的 `ConversionOptions.WithoutAHost` | system temp 只是根目錄；`RecoveryContext.For` 會建立私有 `.aspose-recovery` 子目錄、強制 owner-only 保護，且保護失敗時 fail closed | 符號級 `SuppressMessage` |
| S1075 / S5332 | `Helpers/MhtExternalReferenceScanner.cs` 的 evaluation notice 常數 | URI 是辨識 pinned Aspose evaluation 注入片段的資料，不是連線目標；掃描器只在同程序探測證明 library 會注入，且 decoded active MIME part 的 provenance inspection 證明 caller 未提供完整 notice 時，移除一份完整、精確片段。provenance 僅採用已宣告 boundary（含 RFC transport padding），嚴格處理 transfer encoding / charset，解析不明時 fail closed；caller 的實際 `src`／`href` 仍由 resource scanner 判定，Subject 或附件中的裸 URI 不會被虛構為資源 | 常數範圍 pragma；另有一般與無授權安全回歸測試 |
| CA1068 | `Core/Extension/ExtensionSessionBridge.cs` 的兩個 public unbind API | 重新排列既有公開參數會造成 binary/source compatibility break；內部 WebSocket 方法已改成 cancellation token 最後 | 符號級 `SuppressMessage` |

上述例外若相關契約改變，必須重新評估；尤其是升級 Aspose 版本後，evaluation notice 的來源證明與
精確片段測試都要重跑，不能只因 URI 看起來相同就擴大白名單。

---

## 統計摘要

| 問題類型 | 數量 | 級別 | 主要原因 | 處理方式 |
|----------|------|------|----------|----------|
| AccessToDisposedClosure | 範圍性例外 | Warning | 測試語意與同步 writer contract | .editorconfig / 精確排除 |
| AutoPropertyCanBeMadeGetOnly.Global | 3 | Note | JSON 序列化 | 文件記錄 |
| ClassNeverInstantiated.Global | 1 | Note | 測試類別 | ReSharper disable once |
| CompareOfFloatsByEqualityOperator | 1 | Warning | 整數值比較安全 | ReSharper disable once |
| ConvertToPrimaryConstructor | 1 | Note | 風格選擇 | .editorconfig 排除 |
| MemberCanBePrivate.Global | 6 | Note | 公開 API | 文件記錄 |
| MemberCanBeProtected.Global | 1 | Note | 測試彈性 | 文件記錄 |
| MethodSupportsCancellation | 1 | Note | 複雜度考量 | ReSharper disable once |
| ParameterOnlyUsedForPreconditionCheck.Local | 範圍性例外 | Warning | Assert.All 用法 | .editorconfig 排除 |
| PropertyCanBeMadeInitOnly.Global | 38 | Note | JSON 序列化 | 文件記錄 |
| UnusedMember.Global | 7 | Note | 公開 API | 文件記錄 |
| UnusedMethodReturnValue.Global | 1 | Note | 公開 API | ReSharper disable once |
| UnusedType.Global | 1 | Note | 公開 API | 文件記錄 |
| UseObjectOrCollectionInitializer | 20 | Note | 測試可讀性/誤判 | 文件記錄 |
| UseUtf8StringLiteral | 2 | Note | 誤判 | ReSharper disable/restore |
| MethodHasAsyncOverload | 範圍性例外 | Suggestion | 同步 fixture 順序 | .editorconfig 排除 |
| ClassNeverInstantiated.Local | 1 | Suggestion | JSON 反射建立 DTO | ReSharper disable once |
| SonarCloud / .NET analyzer | 8 類精確例外 | Code smell / Warning | 公開契約、安全不變量、test-only 建議 | NoWarn / 符號級排除 |

> 範圍性規則會隨測試數量變動，因此不再提供容易過期的總數。2026-09-10 重產的
> `report.xml` 在套用上述精確例外後為 0 個 finding。

---

## 如何在檢查時排除這些問題

### 方法 1: 使用 .editorconfig

在專案根目錄的 `.editorconfig` 中添加：

```ini
# 排除特定規則
[*.cs]
dotnet_diagnostic.IDE0017.severity = none  # UseObjectOrCollectionInitializer
dotnet_diagnostic.IDE0059.severity = none  # UnusedVariable (已修復)
```

### 方法 2: 使用 ReSharper/Rider 設定

在 `.DotSettings` 檔案中配置要忽略的規則。

### 方法 3: 使用程式碼註解

對於特定行，可以使用：

```csharp
// ReSharper disable once UnusedMember.Global
public void SomePublicApiMethod() { }
```
