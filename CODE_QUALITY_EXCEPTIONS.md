# Code Quality Exceptions Documentation

本文件記錄 JetBrains InspectCode 報告中被排除修復的問題及其原因。
這些問題經過評估後決定保留，未來進行代碼品質檢查時可參考本文件跳過這些項目。

**最後更新日期**: 2026-01-06
**分析工具**: JetBrains InspectCode 2025.3.0.4

---

## 目錄

1. [AccessToDisposedClosure](#1-accesstodisposedclosure)
2. [AutoPropertyCanBeMadeGetOnly.Global](#2-autopropertycanbemadegetonlyglobal)
3. [MemberCanBePrivate.Global](#3-membercanbeprivateglobal)
4. [MemberCanBeProtected.Global](#4-membercanbeprotectedglobal)
5. [MethodSupportsCancellation](#5-methodsupportscancellation)
6. [OutParameterValueIsAlwaysDiscarded.Local](#6-outparametervalueisalwaysdiscardedlocal)
7. [PropertyCanBeMadeInitOnly.Global](#7-propertycanbemadeinitonlyglobal)
8. [UnusedAutoPropertyAccessor.Global](#8-unusedautopropertyaccessorglobal)
9. [UnusedMember.Global](#9-unusedmemberglobal)
10. [UnusedMember.Local](#10-unusedmemberlocal)
11. [UnusedMethodReturnValue.Global](#11-unusedmethodreturnvalueglobal)
12. [UnusedType.Global](#12-unusedtypeglobal)
13. [UseObjectOrCollectionInitializer](#13-useobjectorollectioninitializer)

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
| **數量** | 2 |
| **訊息** | Auto-property can be made get-only |

### 受影響檔案

| 檔案 | 行號 | 屬性 |
|------|------|------|
| `Core/Security/AuthConfig.cs` | 202 | (JWT 配置屬性) |
| `Core/Security/AuthConfig.cs` | 207 | (JWT 配置屬性) |

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

## 3. MemberCanBePrivate.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 7 |
| **訊息** | Member can be made private |

### 受影響檔案

| 檔案 | 行號 | 成員 |
|------|------|------|
| `Core/Transport/TransportConfig.cs` | 38 | `Host.set` |
| `Core/Transport/TransportConfig.cs` | 28 | `Mode.set` |
| `Core/Transport/TransportConfig.cs` | 33 | `Port.set` |
| `Core/Session/DocumentSession.cs` | 75 | `LastAccessedAt.set` |
| `Core/Tracking/TrackingConfig.cs` | 47 | `WebhookAuthHeader.set` |
| `Core/Tracking/TrackingConfig.cs` | 52 | `WebhookTimeoutSeconds.set` |
| `Core/Session/DocumentSessionManager.cs` | 200 | `GetSession()` |

### 問題描述

成員可以改為 private 可見性。

### 不修復原因

- 這些是公開 API 的一部分，外部可能需要存取
- 配置類屬性需要公開 setter 支援 JSON 反序列化
- `GetSession()` 可能被外部測試或擴展使用
- 降低可見性可能破壞向後相容性

---

## 4. MemberCanBeProtected.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Member can be made protected |

### 受影響檔案

| 檔案 | 行號 | 成員 |
|------|------|------|
| `Tests/Helpers/TestBase.cs` | 19 | `AsposeLibraryType` enum |

### 問題描述

Enum 可以改為 protected 可見性。

### 不修復原因

- 這是測試基類中的公開 enum，供所有測試類使用
- 改為 protected 會限制其在非衍生類中的使用
- 測試程式碼中保持 public 更具彈性

---

## 5. MethodSupportsCancellation

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 2 |
| **訊息** | Method has overload with cancellation support |

### 受影響檔案

| 檔案 | 行號 | 方法 |
|------|------|------|
| `Tests/Core/Security/ApiKeyAuthenticationMiddlewareTests.cs` | 345 | `ReadAsStringAsync` |
| `Tests/Core/Security/ApiKeyAuthenticationMiddlewareTests.cs` | 389 | `ReadAsStringAsync` |

### 問題描述

方法有支援 CancellationToken 的重載可用。

### 不修復原因

- 這是測試代碼，不需要取消支援
- 測試是同步執行的，添加 CancellationToken 會增加不必要的複雜性
- 對測試的可讀性和維護性沒有實質幫助

---

## 6. OutParameterValueIsAlwaysDiscarded.Local

| 項目 | 內容 |
|------|------|
| **級別** | Warning |
| **數量** | 5 |
| **訊息** | Parameter output value is always discarded |

### 受影響檔案

| 檔案 | 行號 | 參數 |
|------|------|------|
| `Tests/Tools/Conversion/ConvertDocumentToolTests.cs` | 107 | `expectedContents` |
| `Tests/Tools/Conversion/ConvertDocumentToolTests.cs` | 168 | `expectedContents` |
| `Tests/Tools/Conversion/ConvertToPdfToolTests.cs` | 36 | `expectedContents` |
| `Tests/Tools/Conversion/ConvertToPdfToolTests.cs` | 89 | `expectedContents` |
| `Tests/Tools/Conversion/ConvertToPdfToolTests.cs` | 143 | `expectedContents` |

### 問題描述

out 參數的輸出值總是被丟棄（使用 `out _`）。

### 不修復原因

- **這是預期行為**，是修復 UnusedVariable 後產生的
- 我們使用 `out _` discard 模式來明確表示不需要輸出值
- 這些輔助方法設計上需要返回 out 參數（給需要它的測試），但有些測試不需要該值
- 這是正確的 C# discard 模式使用

### 範例代碼

```csharp
// 輔助方法設計有 out 參數
private string CreateRichWordDocument(string fileName, out List<string> expectedContents)

// 有些測試需要 expectedContents
var docPath = CreateRichWordDocument("test.docx", out var contents);
foreach (var content in contents) { Assert.Contains(content, result); }

// 有些測試不需要，使用 discard
var docPath = CreateRichWordDocument("test.docx", out _);
```

---

## 7. PropertyCanBeMadeInitOnly.Global

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

## 8. UnusedAutoPropertyAccessor.Global

| 項目 | 內容 |
|------|------|
| **級別** | Warning |
| **數量** | 7 |
| **訊息** | Auto-property accessor is never used |

### 受影響檔案

| 檔案 | 行號 | 屬性 |
|------|------|------|
| `Core/Tracking/TrackingConfig.cs` | 263 | `Error.get` |
| `Core/Session/DocumentSessionManager.cs` | 772 | `EstimatedMemoryMb.get` |
| `Core/Session/DocumentSessionManager.cs` | 767 | `LastAccessedAt.get` |
| `Core/Session/DocumentSessionManager.cs` | 762 | `OpenedAt.get` |
| `Core/Session/TempFileManager.cs` | 605 | `OwnerTenantId.get` |
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

## 9. UnusedMember.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 9 |
| **訊息** | Member is never used |

### 受影響檔案

| 檔案 | 行號 | 成員 |
|------|------|------|
| `Tests/Helpers/ExcelTestBase.cs` | 41 | `AssertCellValue()` |
| `Tests/Helpers/WordTestBase.cs` | 52 | `AssertParagraphExists()` |
| `Tests/Helpers/WordTestBase.cs` | 62 | `AssertParagraphStyle()` |
| `Tests/Helpers/PdfTestBase.cs` | 11 | `IsEvaluationMode()` |
| `Core/Session/DocumentSession.cs` | 226 | `GetDocumentAsync()` |
| `Core/Session/DocumentSessionManager.cs` | 467 | `OnServerShutdown()` |
| `Core/Security/ApiKeyAuthenticationMiddleware.cs` | 423 | `UseApiKeyAuthentication()` |
| `Core/Security/JwtAuthenticationMiddleware.cs` | 537 | `UseJwtAuthentication()` |
| `Core/Tracking/TrackingMiddleware.cs` | 410 | `UseTracking()` |

### 問題描述

成員在專案內部未使用。

### 不修復原因

- 這些是公開 API 的一部分，可能被外部使用者調用
- 測試輔助方法保留供未來測試使用
- 擴展方法 (`UseXxx`) 是設計給外部使用的 API
- 刪除會破壞 API 相容性

---

## 10. UnusedMember.Local

| 項目 | 內容 |
|------|------|
| **級別** | Warning |
| **數量** | 4 |
| **訊息** | Member is never used |

### 受影響檔案

| 檔案 | 行號 | 成員 |
|------|------|------|
| `Tests/Tools/Conversion/ConvertDocumentToolTests.cs` | 95 | `CreateExcelWorkbook()` |
| `Tests/Tools/Conversion/ConvertDocumentToolTests.cs` | 157 | `CreatePowerPointPresentation()` |
| `Tests/Tools/Conversion/ConvertToPdfToolTests.cs` | 77 | `CreateExcelWorkbook()` |
| `Tests/Tools/Conversion/ConvertToPdfToolTests.cs` | 132 | `CreatePowerPointPresentation()` |

### 問題描述

本地私有成員未使用。

### 不修復原因

- 這些是測試輔助方法，可能被未來測試使用
- 保留提供完整的測試基礎設施
- 刪除後如需使用還要重新撰寫
- 與其他已使用的 `CreateRichXxx()` 方法形成完整的測試工具組

---

## 11. UnusedMethodReturnValue.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Return value of method is never used |

### 受影響檔案

| 檔案 | 行號 | 方法 |
|------|------|------|
| `Core/McpServerBuilderExtensions.cs` | 19 | `WithFilteredTools()` |

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

## 12. UnusedType.Global

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 3 |
| **訊息** | Type is never used |

### 受影響檔案

| 檔案 | 行號 | 類型 |
|------|------|------|
| `Core/Security/ApiKeyAuthenticationMiddleware.cs` | 415 | `ApiKeyAuthenticationExtensions` |
| `Core/Security/JwtAuthenticationMiddleware.cs` | 529 | `JwtAuthenticationExtensions` |
| `Core/Tracking/TrackingMiddleware.cs` | 402 | `TrackingExtensions` |

### 問題描述

類型在專案內部未使用。

### 不修復原因

- 公開 API 類型，供外部使用者使用
- 這些是 ASP.NET Core 擴展方法類，用於 `IApplicationBuilder` 配置
- 專案內部使用不同的配置方式，但保留給外部使用者
- 刪除會破壞外部依賴

### 範例代碼

```csharp
// 外部使用者可以這樣配置
app.UseApiKeyAuthentication();
app.UseJwtAuthentication();
app.UseTracking();
```

---

## 13. UseObjectOrCollectionInitializer

| 項目 | 內容 |
|------|------|
| **級別** | Note |
| **數量** | 1 |
| **訊息** | Use object initializer |

### 受影響檔案

| 檔案 | 行號 | 變量 |
|------|------|------|
| `Tests/Tools/Word/WordPropertiesToolTests.cs` | 132 | `doc` |

### 問題描述

建議使用對象初始化器語法。

### 不修復原因

代碼設置的是子對象的屬性，無法在對象初始化器中設置：

```csharp
// 實際代碼 - 設置子對象屬性
var doc = new Document(docPath);
doc.BuiltInDocumentProperties.Title = "Original Title";    // 子對象屬性
doc.BuiltInDocumentProperties.Author = "Original Author";  // 子對象屬性

// 無法使用對象初始化器
var doc = new Document(docPath)
{
    // BuiltInDocumentProperties 是唯讀屬性，無法在此設置其子屬性
};
```

這是 C# 語言限制，不適用初始化器語法。

---

## 統計摘要

| 問題類型 | 數量 | 級別 | 主要原因 |
|----------|------|------|----------|
| AccessToDisposedClosure | 6 | Warning | 測試邏輯需要 |
| AutoPropertyCanBeMadeGetOnly.Global | 2 | Note | JSON 序列化 |
| MemberCanBePrivate.Global | 7 | Note | 公開 API |
| MemberCanBeProtected.Global | 1 | Note | 測試彈性 |
| MethodSupportsCancellation | 2 | Note | 測試不需要 |
| OutParameterValueIsAlwaysDiscarded.Local | 5 | Warning | 預期 discard |
| PropertyCanBeMadeInitOnly.Global | 38 | Note | JSON 序列化 |
| UnusedAutoPropertyAccessor.Global | 7 | Warning | JSON 序列化 / 外部 API |
| UnusedMember.Global | 9 | Note | 公開 API |
| UnusedMember.Local | 4 | Warning | 未來使用 |
| UnusedMethodReturnValue.Global | 1 | Note | 公開 API |
| UnusedType.Global | 3 | Note | 公開 API |
| UseObjectOrCollectionInitializer | 1 | Note | 語言限制 |
| **總計** | **86** | - | - |

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

---

## 維護指南

1. **新增例外時**：請在對應章節添加檔案、行號和原因
2. **移除例外時**：當問題被修復後，從本文件中移除相關記錄
3. **定期審查**：建議每季度審查一次，確認例外是否仍然有效
4. **版本記錄**：重大更新時請更新「最後更新日期」

---

## 變更歷史

| 日期 | 版本 | 變更內容 |
|------|------|----------|
| 2026-01-06 | 1.0.0 | 初始建立文件，記錄 86 個例外項目 |
