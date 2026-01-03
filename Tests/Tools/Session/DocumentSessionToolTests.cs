using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Session;

namespace AsposeMcpServer.Tests.Tools.Session;

public class DocumentSessionToolTests : WordTestBase
{
    private readonly DocumentSessionManager _sessionManager;
    private readonly TempFileManager _tempFileManager;
    private readonly DocumentSessionTool _tool;

    public DocumentSessionToolTests()
    {
        var sessionConfig = new SessionConfig { Enabled = true };
        _sessionManager = new DocumentSessionManager(sessionConfig);
        _tempFileManager = new TempFileManager(sessionConfig);
        _tool = new DocumentSessionTool(_sessionManager, _tempFileManager);
    }

    // Note: This IS the session tool itself, so no separate Session ID Tests region needed

    public override void Dispose()
    {
        _sessionManager.Dispose();
        _tempFileManager.Dispose();
        base.Dispose();
    }

    #region General Tests

    [Fact]
    public void OpenDocument_ShouldReturnSessionId()
    {
        var docPath = CreateWordDocument("test_open.docx");
        var result = _tool.Execute("open", docPath);
        Assert.NotNull(result);
        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.True(json["success"]?.GetValue<bool>());
        Assert.NotNull(json["sessionId"]);
        Assert.StartsWith("sess_", json["sessionId"]?.GetValue<string>());
    }

    [Fact]
    public void ListSessions_AfterOpen_ShouldShowSession()
    {
        var docPath = CreateWordDocument("test_list.docx");
        var openResult = _tool.Execute("open", docPath);
        var openJson = JsonNode.Parse(openResult);
        var sessionId = openJson?["sessionId"]?.GetValue<string>();
        var listResult = _tool.Execute("list");
        var listJson = JsonNode.Parse(listResult);
        Assert.NotNull(listJson);
        Assert.True(listJson["success"]?.GetValue<bool>());
        Assert.Equal(1, listJson["count"]?.GetValue<int>());

        var sessions = listJson["sessions"]?.AsArray();
        Assert.NotNull(sessions);
        Assert.Single(sessions);
        Assert.Equal(sessionId, sessions[0]?["SessionId"]?.GetValue<string>());
    }

    [Fact]
    public void GetStatus_ShouldReturnSessionInfo()
    {
        var docPath = CreateWordDocument("test_status.docx");
        var openResult = _tool.Execute("open", docPath);
        var openJson = JsonNode.Parse(openResult);
        var sessionId = openJson?["sessionId"]?.GetValue<string>();
        var statusResult = _tool.Execute("status", sessionId: sessionId);
        var statusJson = JsonNode.Parse(statusResult);
        Assert.NotNull(statusJson);
        Assert.True(statusJson["success"]?.GetValue<bool>());
        Assert.NotNull(statusJson["session"]);
        Assert.Equal(sessionId, statusJson["session"]?["SessionId"]?.GetValue<string>());
        Assert.Equal("word", statusJson["session"]?["DocumentType"]?.GetValue<string>());
    }

    [Fact]
    public void SaveDocument_ShouldSaveToFile()
    {
        var docPath = CreateWordDocument("test_save.docx");
        var openResult = _tool.Execute("open", docPath);
        var openJson = JsonNode.Parse(openResult);
        var sessionId = openJson?["sessionId"]?.GetValue<string>();

        // Modify the document through session manager directly
        var doc = _sessionManager.GetDocument<Document>(sessionId!);
        var builder = new DocumentBuilder(doc);
        builder.Write("Modified content");
        _sessionManager.MarkDirty(sessionId!);
        var saveResult = _tool.Execute("save", sessionId: sessionId);
        var saveJson = JsonNode.Parse(saveResult);
        Assert.NotNull(saveJson);
        Assert.True(saveJson["success"]?.GetValue<bool>());

        // Verify file was saved
        var savedDoc = new Document(docPath);
        Assert.Contains("Modified content", savedDoc.GetText());
    }

    [Fact]
    public void CloseDocument_ShouldRemoveSession()
    {
        var docPath = CreateWordDocument("test_close.docx");
        var openResult = _tool.Execute("open", docPath);
        var openJson = JsonNode.Parse(openResult);
        var sessionId = openJson?["sessionId"]?.GetValue<string>();

        // Verify session exists
        var listBefore = _tool.Execute("list");
        var listBeforeJson = JsonNode.Parse(listBefore);
        Assert.Equal(1, listBeforeJson?["count"]?.GetValue<int>());
        var closeResult = _tool.Execute("close", sessionId: sessionId);
        var closeJson = JsonNode.Parse(closeResult);
        Assert.NotNull(closeJson);
        Assert.True(closeJson["success"]?.GetValue<bool>());

        // Verify session was removed
        var listAfter = _tool.Execute("list");
        var listAfterJson = JsonNode.Parse(listAfter);
        Assert.Equal(0, listAfterJson?["count"]?.GetValue<int>());
    }

    [Fact]
    public void CloseDocument_WithDiscard_ShouldNotSave()
    {
        var docPath = CreateWordDocument("test_close_discard.docx");
        var openResult = _tool.Execute("open", docPath);
        var openJson = JsonNode.Parse(openResult);
        var sessionId = openJson?["sessionId"]?.GetValue<string>();

        // Modify the document
        var doc = _sessionManager.GetDocument<Document>(sessionId!);
        var builder = new DocumentBuilder(doc);
        builder.Write("Should not be saved");
        _sessionManager.MarkDirty(sessionId!);
        var closeResult = _tool.Execute("close", sessionId: sessionId, discard: true);
        var closeJson = JsonNode.Parse(closeResult);
        Assert.True(closeJson?["success"]?.GetValue<bool>());

        // Verify file was NOT saved with the modification
        var savedDoc = new Document(docPath);
        Assert.DoesNotContain("Should not be saved", savedDoc.GetText());
    }

    [Fact]
    public void OpenDocument_WithReadOnlyMode_ShouldPreventSave()
    {
        var docPath = CreateWordDocument("test_readonly.docx");
        var openResult = _tool.Execute("open", docPath, mode: "readonly");
        var openJson = JsonNode.Parse(openResult);
        var sessionId = openJson?["sessionId"]?.GetValue<string>();
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("save", sessionId: sessionId));
        Assert.Contains("readonly", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void OpenDocument_WithInvalidPath_ShouldThrowException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("open", "nonexistent.docx"));
    }

    [Fact]
    public void GetStatus_WithInvalidSessionId_ShouldThrowException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("status", sessionId: "invalid_session"));
    }

    #endregion
}