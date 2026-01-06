using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Session;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Session;

public class DocumentSessionToolTests : TestBase
{
    private readonly SessionConfig _sessionConfig;
    private readonly DocumentSessionManager _sessionManager;
    private readonly TempFileManager _tempFileManager;
    private readonly DocumentSessionTool _tool;

    public DocumentSessionToolTests()
    {
        _sessionConfig = new SessionConfig
        {
            Enabled = true,
            TempDirectory = Path.Combine(TestDir, "temp"),
            OnDisconnect = DisconnectBehavior.SaveToTemp,
            TempRetentionHours = 24
        };
        Directory.CreateDirectory(_sessionConfig.TempDirectory);
        _sessionManager = new DocumentSessionManager(_sessionConfig);
        _tempFileManager = new TempFileManager(_sessionConfig);
        _tool = new DocumentSessionTool(_sessionManager, _tempFileManager, new StdioSessionIdentityAccessor());
    }

    public override void Dispose()
    {
        _sessionManager.Dispose();
        _tempFileManager.Dispose();
        base.Dispose();
    }

    private string CreateWordDocument(string fileName, string? content = null)
    {
        var filePath = CreateTestFilePath(fileName);
        var doc = new Document();
        if (content != null)
        {
            var builder = new DocumentBuilder(doc);
            builder.Write(content);
        }

        doc.Save(filePath);
        return filePath;
    }

    private string CreateExcelDocument(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(filePath);
        return filePath;
    }

    private string CreatePowerPointDocument(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePdfDocument(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var pdfDoc = new Aspose.Pdf.Document();
        pdfDoc.Pages.Add();
        pdfDoc.Save(filePath);
        return filePath;
    }

    private (string sessionId, JsonNode json) OpenDocument(string path, string mode = "readwrite")
    {
        var result = _tool.Execute("open", path, mode: mode);
        var json = JsonNode.Parse(result)!;
        var sessionId = json["sessionId"]!.GetValue<string>();
        return (sessionId, json);
    }

    private void CreateTempFile(string sessionId, string originalPath, string content = "Temp content")
    {
        var tempDocPath = Path.Combine(_sessionConfig.TempDirectory, $"aspose_session_{sessionId}_20240101120000.docx");
        var tempMetaPath = tempDocPath + ".meta.json";

        var doc = new Document();
        new DocumentBuilder(doc).Write(content);
        doc.Save(tempDocPath);

        var metadata = new
        {
            SessionId = sessionId,
            OriginalPath = originalPath,
            TempPath = tempDocPath,
            DocumentType = "Word",
            SavedAt = DateTime.UtcNow,
            PromptOnReconnect = false
        };
        File.WriteAllText(tempMetaPath, JsonSerializer.Serialize(metadata));
    }

    #region Open

    [Theory]
    [InlineData(".docx", "word")]
    [InlineData(".xlsx", "excel")]
    [InlineData(".pptx", "powerpoint")]
    [InlineData(".pdf", "pdf")]
    public void Open_DifferentDocumentTypes_ShouldReturnCorrectDocumentType(string extension, string expectedType)
    {
        var docPath = extension switch
        {
            ".docx" => CreateWordDocument($"test_open{extension}"),
            ".xlsx" => CreateExcelDocument($"test_open{extension}"),
            ".pptx" => CreatePowerPointDocument($"test_open{extension}"),
            ".pdf" => CreatePdfDocument($"test_open{extension}"),
            _ => throw new ArgumentException($"Unsupported extension: {extension}")
        };

        var (sessionId, json) = OpenDocument(docPath);

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.StartsWith("sess_", sessionId);

        var statusResult = _tool.Execute("status", sessionId: sessionId);
        var statusJson = JsonNode.Parse(statusResult)!;
        Assert.Equal(expectedType, statusJson["session"]!["DocumentType"]!.GetValue<string>());
    }

    [Fact]
    public void Open_WithReadOnlyMode_ShouldPreventSave()
    {
        var docPath = CreateWordDocument("test_readonly.docx");
        var (sessionId, _) = OpenDocument(docPath, "readonly");

        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("save", sessionId: sessionId));
        Assert.Contains("readonly", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("OPEN")]
    [InlineData("Open")]
    [InlineData("open")]
    public void Open_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var result = _tool.Execute(operation, docPath);
        var json = JsonNode.Parse(result)!;
        Assert.True(json["success"]!.GetValue<bool>());
    }

    #endregion

    #region Save

    [Fact]
    public void Save_Word_ShouldSaveModifications()
    {
        var docPath = CreateWordDocument("test_save_word.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var doc = _sessionManager.GetDocument<Document>(sessionId);
        new DocumentBuilder(doc).Write("Modified content");
        _sessionManager.MarkDirty(sessionId);

        var saveResult = _tool.Execute("save", sessionId: sessionId);
        var saveJson = JsonNode.Parse(saveResult)!;
        Assert.True(saveJson["success"]!.GetValue<bool>());

        var savedDoc = new Document(docPath);
        Assert.Contains("Modified content", savedDoc.GetText());
    }

    [Fact]
    public void Save_Excel_ShouldSaveModifications()
    {
        var docPath = CreateExcelDocument("test_save_excel.xlsx");
        var (sessionId, _) = OpenDocument(docPath);

        var workbook = _sessionManager.GetDocument<Workbook>(sessionId);
        workbook.Worksheets[0].Cells["B1"].Value = "Modified";
        _sessionManager.MarkDirty(sessionId);

        var saveResult = _tool.Execute("save", sessionId: sessionId);
        var saveJson = JsonNode.Parse(saveResult)!;
        Assert.True(saveJson["success"]!.GetValue<bool>());

        var savedWorkbook = new Workbook(docPath);
        Assert.Equal("Modified", savedWorkbook.Worksheets[0].Cells["B1"].StringValue);
    }

    [Fact]
    public void Save_WithOutputPath_ShouldSaveToNewLocation()
    {
        var docPath = CreateWordDocument("test_save_output.docx");
        var outputPath = CreateTestFilePath("test_save_new.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var doc = _sessionManager.GetDocument<Document>(sessionId);
        new DocumentBuilder(doc).Write("New content");
        _sessionManager.MarkDirty(sessionId);

        var saveResult = _tool.Execute("save", sessionId: sessionId, outputPath: outputPath);
        var saveJson = JsonNode.Parse(saveResult)!;

        Assert.True(saveJson["success"]!.GetValue<bool>());
        Assert.Contains(outputPath, saveJson["message"]!.GetValue<string>());
        Assert.True(File.Exists(outputPath));

        var savedDoc = new Document(outputPath);
        Assert.Contains("New content", savedDoc.GetText());
    }

    #endregion

    #region Close

    [Fact]
    public void Close_ShouldRemoveSession()
    {
        var docPath = CreateWordDocument("test_close.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var listBefore = JsonNode.Parse(_tool.Execute("list"))!;
        Assert.Equal(1, listBefore["count"]!.GetValue<int>());

        var closeResult = _tool.Execute("close", sessionId: sessionId);
        var closeJson = JsonNode.Parse(closeResult)!;
        Assert.True(closeJson["success"]!.GetValue<bool>());

        var listAfter = JsonNode.Parse(_tool.Execute("list"))!;
        Assert.Equal(0, listAfter["count"]!.GetValue<int>());
    }

    [Fact]
    public void Close_WithoutDiscard_ShouldAutoSave()
    {
        var docPath = CreateWordDocument("test_close_autosave.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var doc = _sessionManager.GetDocument<Document>(sessionId);
        new DocumentBuilder(doc).Write("Auto-saved content");
        _sessionManager.MarkDirty(sessionId);

        var closeResult = _tool.Execute("close", sessionId: sessionId, discard: false);
        var closeJson = JsonNode.Parse(closeResult)!;
        Assert.True(closeJson["success"]!.GetValue<bool>());
        Assert.Contains("changes saved", closeJson["message"]!.GetValue<string>());

        var savedDoc = new Document(docPath);
        Assert.Contains("Auto-saved content", savedDoc.GetText());
    }

    [Fact]
    public void Close_WithDiscard_Word_ShouldNotSave()
    {
        var docPath = CreateWordDocument("test_close_discard_word.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var doc = _sessionManager.GetDocument<Document>(sessionId);
        new DocumentBuilder(doc).Write("Should not be saved");
        _sessionManager.MarkDirty(sessionId);

        var closeResult = _tool.Execute("close", sessionId: sessionId, discard: true);
        var closeJson = JsonNode.Parse(closeResult)!;
        Assert.True(closeJson["success"]!.GetValue<bool>());

        var savedDoc = new Document(docPath);
        Assert.DoesNotContain("Should not be saved", savedDoc.GetText());
    }

    [Fact]
    public void Close_WithDiscard_Excel_ShouldNotSave()
    {
        var docPath = CreateExcelDocument("test_close_discard_excel.xlsx");
        var (sessionId, _) = OpenDocument(docPath);

        var workbook = _sessionManager.GetDocument<Workbook>(sessionId);
        workbook.Worksheets[0].Cells["C1"].Value = "Should not be saved";
        _sessionManager.MarkDirty(sessionId);

        var closeResult = _tool.Execute("close", sessionId: sessionId, discard: true);
        var closeJson = JsonNode.Parse(closeResult)!;
        Assert.True(closeJson["success"]!.GetValue<bool>());

        var savedWorkbook = new Workbook(docPath);
        Assert.NotEqual("Should not be saved", savedWorkbook.Worksheets[0].Cells["C1"].StringValue);
    }

    #endregion

    #region List & Status

    [Fact]
    public void List_AfterOpen_ShouldShowSession()
    {
        var docPath = CreateWordDocument("test_list.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var listResult = _tool.Execute("list");
        var listJson = JsonNode.Parse(listResult)!;

        Assert.True(listJson["success"]!.GetValue<bool>());
        Assert.Equal(1, listJson["count"]!.GetValue<int>());

        var sessions = listJson["sessions"]!.AsArray();
        Assert.Single(sessions);
        Assert.Equal(sessionId, sessions[0]!["SessionId"]!.GetValue<string>());
    }

    [Fact]
    public void List_MultipleDifferentTypes_ShouldShowAll()
    {
        var wordPath = CreateWordDocument("test_multi.docx");
        var excelPath = CreateExcelDocument("test_multi.xlsx");
        var pptPath = CreatePowerPointDocument("test_multi.pptx");

        _tool.Execute("open", wordPath);
        _tool.Execute("open", excelPath);
        _tool.Execute("open", pptPath);

        var listResult = _tool.Execute("list");
        var listJson = JsonNode.Parse(listResult)!;

        Assert.True(listJson["success"]!.GetValue<bool>());
        Assert.Equal(3, listJson["count"]!.GetValue<int>());

        var sessions = listJson["sessions"]!.AsArray();
        var docTypes = sessions.Select(s => s!["DocumentType"]!.GetValue<string>()).ToList();
        Assert.Contains("word", docTypes);
        Assert.Contains("excel", docTypes);
        Assert.Contains("powerpoint", docTypes);
    }

    [Fact]
    public void Status_ShouldReturnSessionInfo()
    {
        var docPath = CreateWordDocument("test_status.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var statusResult = _tool.Execute("status", sessionId: sessionId);
        var statusJson = JsonNode.Parse(statusResult)!;

        Assert.True(statusJson["success"]!.GetValue<bool>());
        Assert.NotNull(statusJson["session"]);
        Assert.Equal("word", statusJson["session"]!["DocumentType"]!.GetValue<string>());
    }

    #endregion

    #region TempFile

    [Fact]
    public void ListTemp_WhenNoTempFiles_ShouldReturnEmptyList()
    {
        var result = _tool.Execute("list_temp");
        var json = JsonNode.Parse(result)!;

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.Equal(0, json["count"]!.GetValue<int>());
        Assert.NotNull(json["files"]!.AsArray());
    }

    [Fact]
    public void TempStats_WhenNoTempFiles_ShouldReturnZeroStats()
    {
        var result = _tool.Execute("temp_stats");
        var json = JsonNode.Parse(result)!;

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.Equal(0, json["TotalCount"]!.GetValue<int>());
        Assert.Equal(0, json["ExpiredCount"]!.GetValue<int>());
    }

    [Fact]
    public void Cleanup_WhenNoTempFiles_ShouldSucceed()
    {
        var result = _tool.Execute("cleanup");
        var json = JsonNode.Parse(result)!;

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.Equal(0, json["DeletedCount"]!.GetValue<int>());
    }

    [Fact]
    public void TempFileOperations_FullWorkflow()
    {
        var sessionId = "sess_test123456";
        var originalPath = CreateWordDocument("original_temp_test.docx");
        CreateTempFile(sessionId, originalPath, "Temp file content");

        var listResult = _tool.Execute("list_temp");
        var listJson = JsonNode.Parse(listResult)!;
        Assert.True(listJson["success"]!.GetValue<bool>());
        Assert.Equal(1, listJson["count"]!.GetValue<int>());
        Assert.Equal(sessionId, listJson["files"]!.AsArray()[0]!["SessionId"]!.GetValue<string>());

        var statsResult = _tool.Execute("temp_stats");
        var statsJson = JsonNode.Parse(statsResult)!;
        Assert.True(statsJson["success"]!.GetValue<bool>());
        Assert.Equal(1, statsJson["TotalCount"]!.GetValue<int>());

        var recoverPath = CreateTestFilePath("recovered_temp_test.docx");
        var recoverResult = _tool.Execute("recover", sessionId: sessionId, outputPath: recoverPath,
            deleteAfterRecover: false);
        var recoverJson = JsonNode.Parse(recoverResult)!;
        Assert.True(recoverJson["Success"]!.GetValue<bool>());
        Assert.True(File.Exists(recoverPath));
        Assert.Contains("Temp file content", new Document(recoverPath).GetText());

        var deleteResult = _tool.Execute("delete_temp", sessionId: sessionId);
        var deleteJson = JsonNode.Parse(deleteResult)!;
        Assert.True(deleteJson["success"]!.GetValue<bool>());
    }

    [Fact]
    public void Recover_WithDeleteAfterRecover_ShouldDeleteTempFile()
    {
        var sessionId = "sess_delafter123";
        var originalPath = CreateWordDocument("original_del_after.docx");
        var tempDocPath = Path.Combine(_sessionConfig.TempDirectory, $"aspose_session_{sessionId}_20240101120000.docx");
        var tempMetaPath = tempDocPath + ".meta.json";
        CreateTempFile(sessionId, originalPath, "Delete after recover content");

        var recoverPath = CreateTestFilePath("recovered_del_after.docx");
        var result = _tool.Execute("recover", sessionId: sessionId, outputPath: recoverPath, deleteAfterRecover: true);
        var json = JsonNode.Parse(result)!;

        Assert.True(json["Success"]!.GetValue<bool>());
        Assert.False(File.Exists(tempDocPath));
        Assert.False(File.Exists(tempMetaPath));
        Assert.True(File.Exists(recoverPath));
    }

    [Fact]
    public void Recover_ToOriginalPath_ShouldUseOriginalPath()
    {
        var sessionId = "sess_origpath123";
        var originalPath = CreateTestFilePath("original_path_test.docx");
        CreateTempFile(sessionId, originalPath, "Recover to original");

        var result = _tool.Execute("recover", sessionId: sessionId, outputPath: null, deleteAfterRecover: true);
        var json = JsonNode.Parse(result)!;

        Assert.True(json["Success"]!.GetValue<bool>());
        Assert.Equal(originalPath, json["RecoveredPath"]!.GetValue<string>());
        Assert.True(File.Exists(originalPath));
        Assert.Contains("Recover to original", new Document(originalPath).GetText());
    }

    [Fact]
    public void Cleanup_ShouldDeleteExpiredFiles()
    {
        var expiredConfig = new SessionConfig
        {
            Enabled = true,
            TempDirectory = _sessionConfig.TempDirectory,
            TempRetentionHours = 0
        };
        var expiredTempManager = new TempFileManager(expiredConfig);
        var expiredTool =
            new DocumentSessionTool(_sessionManager, expiredTempManager, new StdioSessionIdentityAccessor());

        var sessionId = "sess_expired123";
        var originalPath = CreateWordDocument("original_expired.docx");
        var tempDocPath = Path.Combine(_sessionConfig.TempDirectory, $"aspose_session_{sessionId}_20240101120000.docx");
        var tempMetaPath = tempDocPath + ".meta.json";

        var doc = new Document();
        doc.Save(tempDocPath);

        var metadata = new
        {
            SessionId = sessionId,
            OriginalPath = originalPath,
            TempPath = tempDocPath,
            DocumentType = "Word",
            SavedAt = DateTime.UtcNow.AddHours(-25),
            PromptOnReconnect = false
        };
        File.WriteAllText(tempMetaPath, JsonSerializer.Serialize(metadata));

        var result = expiredTool.Execute("cleanup");
        var json = JsonNode.Parse(result)!;

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.True(json["DeletedCount"]!.GetValue<int>() >= 1);
        Assert.False(File.Exists(tempDocPath));
        Assert.False(File.Exists(tempMetaPath));

        expiredTempManager.Dispose();
    }

    [Fact]
    public void TempStats_WithMultipleFiles_ShouldReturnCorrectStatistics()
    {
        for (var i = 0; i < 3; i++)
        {
            var sessionId = $"sess_stats_{i}";
            var originalPath = CreateWordDocument($"original_stats_{i}.docx");
            CreateTempFile(sessionId, originalPath, $"Stats test content {i}");
        }

        var result = _tool.Execute("temp_stats");
        var json = JsonNode.Parse(result)!;

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.Equal(3, json["TotalCount"]!.GetValue<int>());
        Assert.True(json["TotalSizeMb"]!.GetValue<double>() > 0);
        Assert.Equal(24, json["retentionHours"]!.GetValue<int>());
    }

    [Fact]
    public void DeleteTemp_WithNonExistentSessionId_ShouldReturnFailure()
    {
        var result = _tool.Execute("delete_temp", sessionId: "nonexistent_session");
        var json = JsonNode.Parse(result)!;

        Assert.False(json["success"]!.GetValue<bool>());
        Assert.Contains("not found", json["message"]!.GetValue<string>());
    }

    [Fact]
    public void Recover_WithNonExistentSessionId_ShouldReturnFailure()
    {
        var result = _tool.Execute("recover", sessionId: "nonexistent_session");
        var json = JsonNode.Parse(result)!;

        Assert.False(json["Success"]!.GetValue<bool>());
        Assert.NotNull(json["ErrorMessage"]!.GetValue<string>());
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown_operation"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Open_WithInvalidPath_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() => _tool.Execute("open", "nonexistent.docx"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Open_WithNullOrEmptyPath_ShouldThrowArgumentException(string? path)
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("open", path));
        Assert.Contains("path is required", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Save_WithNullOrEmptySessionId_ShouldThrowArgumentException(string? sessionId)
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("save", sessionId: sessionId));
        Assert.Contains("sessionId is required", ex.Message);
    }

    [Fact]
    public void Save_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("save", sessionId: "invalid_session"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Close_WithNullOrEmptySessionId_ShouldThrowArgumentException(string? sessionId)
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("close", sessionId: sessionId));
        Assert.Contains("sessionId is required", ex.Message);
    }

    [Fact]
    public void Close_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("close", sessionId: "invalid_session"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Status_WithNullOrEmptySessionId_ShouldThrowArgumentException(string? sessionId)
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("status", sessionId: sessionId));
        Assert.Contains("sessionId is required", ex.Message);
    }

    [Fact]
    public void Status_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("status", sessionId: "invalid_session"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void DeleteTemp_WithNullOrEmptySessionId_ShouldThrowArgumentException(string? sessionId)
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete_temp", sessionId: sessionId));
        Assert.Contains("sessionId is required", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Recover_WithNullOrEmptySessionId_ShouldThrowArgumentException(string? sessionId)
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("recover", sessionId: sessionId));
        Assert.Contains("sessionId is required", ex.Message);
    }

    #endregion
}