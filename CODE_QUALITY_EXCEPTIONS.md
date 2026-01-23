# Code Quality Exceptions Documentation

本文件記錄 JetBrains InspectCode 報告中被排除修復的問題及其原因。
這些問題經過評估後決定保留，未來進行代碼品質檢查時可參考本文件跳過這些項目。

**最後更新日期**: 2026-01-23
**分析工具**: JetBrains InspectCode 2025.3.0.4

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
11. [UnusedAutoPropertyAccessor.Global](#11-unusedautopropertyaccessorglobal)
12. [UnusedMember.Global](#12-unusedmemberglobal)
13. [UnusedMethodReturnValue.Global](#13-unusedmethodreturnvalueglobal)
14. [UnusedType.Global](#14-unusedtypeglobal)
15. [UseObjectOrCollectionInitializer](#15-useobjectorollectioninitializer)
16. [UseUtf8StringLiteral](#16-useutf8stringliteral)

---

## 1. AccessToDisposedClosure

| 項目 | 內容 |
|------|------|
| **級別** | Warning |
| **數量** | 6 |
| **訊息** | Captured variable is disposed in the outer scope |

### 受影響檔案

| 檔案 | 行號 | 變量 |
|------|------|------|
| `Tests/Core/Helpers/AsposeHelperTests.cs` | 21 | `workbook` |
| `Tests/Core/Helpers/AsposeHelperTests.cs` | 251 | `presentation` |
| `Tests/Core/Session/DocumentContextTests.cs` | 265 | `context` |
| `Tests/Core/Session/DocumentSessionManagerTests.cs` | 169 | `manager` |
| `Tests/Core/Session/DocumentSessionManagerTests.cs` | 408 | `manager` |
| `Tests/Core/Session/DocumentSessionManagerTests.cs` | 423 | `manager` |

### 問題描述

在 lambda 或匿名方法中捕獲的變量在外部作用域被 dispose。

### 不修復原因

- 這是測試代碼中的模式，用於測試 `Assert.Throws` 等異常處理場景
- 這些是故意設計的測試場景，用於驗證對象 dispose 後的行為
- 修改可能導致測試無法正確驗證預期行為
- 測試需要驗證「已釋放對象被存取時應拋出異常」的行為

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

| 檔案 | 行號 | 屬性 |
|------|------|------|
| `Core/Security/JwtConfig.cs` | 73 | `ClientSecret` (JWT 配置屬性) |
| `Core/Security/JwtConfig.cs` | 79 | `CustomEndpoint` (JWT 配置屬性) |
| `Core/Security/ApiKeyConfig.cs` | 47 | `CustomEndpoint` (API Key 配置屬性) |

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

| 檔案 | 行號 | 類別 |
|------|------|------|
| `Tests/Core/Handlers/HandlerRegistryAutoDiscoveryTests.cs` | 116 | `DifferentContextDocument` |

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

| 檔案 | 行號 | 說明 |
|------|------|------|
| `Tests/Handlers/Excel/DataOperations/GetContentHandlerTests.cs` | 43 | 整數值 100 的浮點比較 |

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

| 檔案 | 行號 | 類別 |
|------|------|------|
| `Tests/Core/Handlers/HandlerRegistryAutoDiscoveryTests.cs` | 160 | `NoParameterlessCtorHandler` |

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

| 檔案 | 行號 | 成員 |
|------|------|------|
| `Core/Transport/TransportConfig.cs` | 28 | `Mode.set` |
| `Core/Transport/TransportConfig.cs` | 33 | `Port.set` |
| `Core/Transport/TransportConfig.cs` | 38 | `Host.set` |
| `Core/Session/DocumentSession.cs` | 75 | `LastAccessedAt.set` |
| `Core/Tracking/TrackingConfig.cs` | 47 | `WebhookAuthHeader.set` |
| `Core/Tracking/TrackingConfig.cs` | 52 | `WebhookTimeoutSeconds.set` |

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

| 檔案 | 行號 | 成員 |
|------|------|------|
| `Tests/Infrastructure/TestBase.cs` | 526 | `AsposeLibraryType` enum |

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

| 檔案 | 行號 | 方法 | 備註 |
|------|------|------|------|
| `Core/Transport/WebSocketConnectionHandler.cs` | 108 | `WaitAsync` | 生產代碼 |

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

| 檔案 | 行號 | 參數 |
|------|------|------|
| `Tests/Helpers/PowerPoint/PptLayoutHelperTests.cs` | 21 | `item` (Assert.All lambda 參數) |
| `Tests/Handlers/Word/Text/SearchWordTextHandlerTests.cs` | 345 | `m` (Assert.All lambda 參數) |

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

## 11. UnusedAutoPropertyAccessor.Global

| 項目 | 內容 |
|------|------|
| **級別** | Warning |
| **數量** | 7 |
| **訊息** | Auto-property accessor is never used |

### 受影響檔案

| 檔案 | 行號 | 屬性 |
|------|------|------|
| `Core/Tracking/TrackingConfig.cs` | 264 | `Error.get` |
| `Core/Session/DocumentSessionManager.cs` | 771 | `EstimatedMemoryMb.get` |
| `Core/Session/DocumentSessionManager.cs` | 766 | `LastAccessedAt.get` |
| `Core/Session/DocumentSessionManager.cs` | 761 | `OpenedAt.get` |
| `Core/Session/TempFileManager.cs` | 605 | `OwnerGroupId.get` |
| `Core/Session/TempFileManager.cs` | 610 | `OwnerUserId.get` |
| `Core/Session/TempFileManager.cs` | 575 | `TempPath.get` |

### 問題描述

自動屬性的 getter 或 setter 在專案內部未被使用。

### 不修復原因

- Setter 用於 JSON 反序列化
- 這些屬性可能被外部代碼或未來功能使用
- 移除會影響序列化/反序列化行為
- 這些是狀態記錄屬性，提供給外部診斷使用

---

## 12. UnusedMember.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 7 |
| **訊息** | Member is never used |

### 受影響檔案

| 檔案 | 行號 | 成員 |
|------|------|------|
| `Tests/Infrastructure/ExcelTestBase.cs` | 41 | `AssertCellValue()` |
| `Tests/Infrastructure/WordTestBase.cs` | 52 | `AssertParagraphExists()` |
| `Tests/Infrastructure/WordTestBase.cs` | 62 | `AssertParagraphStyle()` |
| `Tests/Infrastructure/PdfTestBase.cs` | 11 | `IsEvaluationMode()` |
| `Core/Session/DocumentSession.cs` | 226 | `GetDocumentAsync()` |
| `Core/Session/DocumentSessionManager.cs` | 467 | `OnServerShutdown()` |
| `Core/Tracking/TrackingExtensions.cs` | 14 | `UseTracking()` |

### 問題描述

成員在專案內部未使用。

### 不修復原因

- 這些是公開 API 的一部分，可能被外部使用者調用
- 測試輔助方法保留供未來測試使用
- 擴展方法 (`UseXxx`) 是設計給外部使用的 API
- 刪除會破壞 API 相容性

---

## 13. UnusedMethodReturnValue.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Return value of method is never used |
| **處理方式** | 已添加 `// ReSharper disable once` 註解 |

### 受影響檔案

| 檔案 | 行號 | 方法 |
|------|------|------|
| `Core/McpServerBuilderExtensions.cs` | 20 | `WithFilteredTools()` |

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

## 14. UnusedType.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Type is never used |

### 受影響檔案

| 檔案 | 行號 | 類型 |
|------|------|------|
| `Core/Tracking/TrackingExtensions.cs` | 6 | `TrackingExtensions` |

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

## 15. UseObjectOrCollectionInitializer

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 26 |
| **訊息** | Use object initializer |

### 受影響檔案

#### 測試檔案 - AuthConfigTests.cs (20 個)

這些是測試程式碼中對 `AuthConfig` 嵌套屬性的設定：

| 檔案 | 行號 |
|------|------|
| `Tests/Core/Security/AuthConfigTests.cs` | 629, 641, 653, 665, 677, 689, 701, 715, 730, 743, 755, 767, 779, 791, 805, 820, 830, 840, 851 |

#### 測試檔案 - GetParagraphFormatWordHandlerTests.cs (5 個)

| 檔案 | 行號 | 說明 |
|------|------|------|
| `Tests/Handlers/Word/Paragraph/GetParagraphFormatWordHandlerTests.cs` | 313, 338, 429, 456, 479 | 誤判 |

#### 測試檔案 - AddTableOfContentsWordHandlerTests.cs (1 個)

| 檔案 | 行號 | 說明 |
|------|------|------|
| `Tests/Handlers/Word/Reference/AddTableOfContentsWordHandlerTests.cs` | 26 | 誤判 |

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

## 16. UseUtf8StringLiteral

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 2 |
| **訊息** | Collection expression can be converted to a UTF-8 string literal |
| **處理方式** | 已添加 `// ReSharper disable` / `// ReSharper restore` 註解對 (行 72-82) |

### 受影響檔案

| 檔案 | 行號 | 說明 |
|------|------|------|
| `Tests/Core/ShapeDetailProviders/PictureFrameDetailProviderTests.cs` | 73-81 | 誤判 (PNG 二進制數據) |

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

## 統計摘要

| 問題類型 | 數量 | 級別 | 主要原因 | 處理方式 |
|----------|------|------|----------|----------|
| AccessToDisposedClosure | 6 | Warning | 測試邏輯需要 | 文件記錄 |
| AutoPropertyCanBeMadeGetOnly.Global | 3 | Note | JSON 序列化 | 文件記錄 |
| ClassNeverInstantiated.Global | 1 | Note | 測試類別 | ReSharper disable once |
| CompareOfFloatsByEqualityOperator | 1 | Warning | 整數值比較安全 | ReSharper disable once |
| ConvertToPrimaryConstructor | 1 | Note | 風格選擇 | .editorconfig 排除 |
| MemberCanBePrivate.Global | 6 | Note | 公開 API | 文件記錄 |
| MemberCanBeProtected.Global | 1 | Note | 測試彈性 | 文件記錄 |
| MethodSupportsCancellation | 1 | Note | 複雜度考量 | ReSharper disable once |
| ParameterOnlyUsedForPreconditionCheck.Local | 2 | Warning | Assert.All 用法 | ReSharper disable once |
| PropertyCanBeMadeInitOnly.Global | 38 | Note | JSON 序列化 | 文件記錄 |
| UnusedAutoPropertyAccessor.Global | 7 | Warning | JSON 序列化 / 外部 API | 文件記錄 |
| UnusedMember.Global | 7 | Note | 公開 API | 文件記錄 |
| UnusedMethodReturnValue.Global | 1 | Note | 公開 API | ReSharper disable once |
| UnusedType.Global | 1 | Note | 公開 API | 文件記錄 |
| UseObjectOrCollectionInitializer | 26 | Note | 測試可讀性/誤判 | 文件記錄 |
| UseUtf8StringLiteral | 2 | Note | 誤判 | ReSharper disable/restore |
| **總計** | **104** | - | - | - |

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
