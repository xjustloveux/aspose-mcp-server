using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Session;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Session;

/// <summary>
///     Integration tests for DocumentSessionTool.
///     Focuses on core session operations, file I/O, and operation routing.
///     Detailed parameter validation tests are in Handler tests.
/// </summary>
public class DocumentSessionToolTests : TestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly TempFileManager _tempFileManager;
    private readonly DocumentSessionTool _tool;

    public DocumentSessionToolTests()
    {
        var sessionConfig = new SessionConfig
        {
            Enabled = true,
            TempDirectory = Path.Combine(TestDir, "temp"),
            OnDisconnect = DisconnectBehavior.SaveToTemp,
            TempRetentionHours = 24
        };
        Directory.CreateDirectory(sessionConfig.TempDirectory);
        _sessionManager = new DocumentSessionManager(sessionConfig);
        _tempFileManager = new TempFileManager(sessionConfig);
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
        if (content != null) new DocumentBuilder(doc).Write(content);
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

    private (string sessionId, JsonNode json) OpenDocument(string path, string mode = "readwrite")
    {
        var result = _tool.Execute("open", path, mode: mode);
        var json = JsonNode.Parse(result)!;
        var sessionId = json["sessionId"]!.GetValue<string>();
        return (sessionId, json);
    }

    #region File I/O Smoke Tests

    [Theory]
    [InlineData(".docx", "word")]
    [InlineData(".xlsx", "excel")]
    [InlineData(".pptx", "powerpoint")]
    public void Open_DifferentDocumentTypes_ShouldReturnCorrectDocumentType(string extension, string expectedType)
    {
        var docPath = extension switch
        {
            ".docx" => CreateWordDocument($"test_open{extension}"),
            ".xlsx" => CreateExcelDocument($"test_open{extension}"),
            ".pptx" => CreatePowerPointDocument($"test_open{extension}"),
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
    public void Save_Word_ShouldPersistModifications()
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
        Assert.True(File.Exists(outputPath));
    }

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
    public void List_AfterOpen_ShouldShowSession()
    {
        var docPath = CreateWordDocument("test_list.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var listResult = _tool.Execute("list");
        var listJson = JsonNode.Parse(listResult)!;

        Assert.True(listJson["success"]!.GetValue<bool>());
        Assert.Equal(1, listJson["count"]!.GetValue<int>());
        var sessions = listJson["sessions"]!.AsArray();
        Assert.Equal(sessionId, sessions[0]!["SessionId"]!.GetValue<string>());
    }

    [Fact]
    public void Status_ShouldReturnSessionInfo()
    {
        var docPath = CreateWordDocument("test_status.docx");
        var (sessionId, _) = OpenDocument(docPath);

        var statusResult = _tool.Execute("status", sessionId: sessionId);
        var statusJson = JsonNode.Parse(statusResult)!;

        Assert.True(statusJson["success"]!.GetValue<bool>());
        Assert.Equal("word", statusJson["session"]!["DocumentType"]!.GetValue<string>());
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("OPEN")]
    [InlineData("Open")]
    [InlineData("open")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var result = _tool.Execute(operation, docPath);
        var json = JsonNode.Parse(result)!;
        Assert.True(json["success"]!.GetValue<bool>());
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown_operation"));
        Assert.Contains("Unknown operation", ex.Message);
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

    #endregion

    #region TempFile Operations

    [Fact]
    public void ListTemp_WhenNoTempFiles_ShouldReturnEmptyList()
    {
        var result = _tool.Execute("list_temp");
        var json = JsonNode.Parse(result)!;

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.Equal(0, json["count"]!.GetValue<int>());
    }

    [Fact]
    public void TempStats_WhenNoTempFiles_ShouldReturnZeroStats()
    {
        var result = _tool.Execute("temp_stats");
        var json = JsonNode.Parse(result)!;

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.Equal(0, json["TotalCount"]!.GetValue<int>());
    }

    [Fact]
    public void Cleanup_WhenNoTempFiles_ShouldSucceed()
    {
        var result = _tool.Execute("cleanup");
        var json = JsonNode.Parse(result)!;

        Assert.True(json["success"]!.GetValue<bool>());
        Assert.Equal(0, json["DeletedCount"]!.GetValue<int>());
    }

    #endregion

    #region Exception Handling

    [Fact]
    public void Open_WithInvalidPath_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() => _tool.Execute("open", "nonexistent.docx"));
    }

    [Fact]
    public void Save_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("save", sessionId: "invalid_session"));
    }

    [Fact]
    public void Close_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("close", sessionId: "invalid_session"));
    }

    [Fact]
    public void Status_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("status", sessionId: "invalid_session"));
    }

    #endregion
}
