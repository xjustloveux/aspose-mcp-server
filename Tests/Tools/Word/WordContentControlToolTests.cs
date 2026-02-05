using Aspose.Words;
using Aspose.Words.Markup;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.ContentControl;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordContentControlTool.
/// </summary>
public class WordContentControlToolTests : WordTestBase
{
    private readonly WordContentControlTool _tool;

    public WordContentControlToolTests()
    {
        _tool = new WordContentControlTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddContentControl_ShouldAddAndPersistToFile()
    {
        var docPath = CreateWordDocumentWithContent("test_cc_add.docx", "Test content");
        var outputPath = CreateTestFilePath("test_cc_add_output.docx");
        _tool.Execute("add", docPath, outputPath: outputPath, type: "PlainText", tag: "testCC", title: "Test CC");
        var doc = new Document(outputPath);
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        Assert.True(sdtNodes.Count > 0);
        var sdt = (StructuredDocumentTag)sdtNodes[0];
        Assert.Equal("testCC", sdt.Tag);
    }

    [Fact]
    public void GetContentControls_ShouldReturnFromFile()
    {
        var docPath = CreateWordDocument("test_cc_get.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "existing" };
        builder.InsertNode(sdt);
        doc.Save(docPath);

        var result = _tool.Execute("get", docPath);
        var data = GetResultData<GetContentControlsResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.ContentControls, cc => cc.Tag == "existing");
    }

    [Fact]
    public void DeleteContentControl_ShouldDeleteAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_cc_delete.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "toDelete" };
        builder.InsertNode(sdt);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_cc_delete_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, tag: "toDelete");
        var resultDoc = new Document(outputPath);
        var sdtNodes = resultDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        Assert.Equal(0, sdtNodes.Count);
    }

    [Fact]
    public void SetValue_ShouldSetValueAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_cc_setval.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "valueTag" };
        builder.InsertNode(sdt);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_cc_setval_output.docx");
        _tool.Execute("set_value", docPath, outputPath: outputPath, tag: "valueTag", value: "Hello");
        var resultDoc = new Document(outputPath);
        var resultSdt =
            (StructuredDocumentTag)resultDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.Contains("Hello", resultSdt.GetText());
    }

    [Fact]
    public void EditContentControl_ShouldEditAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_cc_edit.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Tag = "editTag",
            Title = "Old Title"
        };
        builder.InsertNode(sdt);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_cc_edit_output.docx");
        _tool.Execute("edit", docPath, outputPath: outputPath, index: 0, newTitle: "New Title");
        var resultDoc = new Document(outputPath);
        var resultSdt =
            (StructuredDocumentTag)resultDoc.GetChildNodes(NodeType.StructuredDocumentTag, true)[0];
        Assert.Equal("New Title", resultSdt.Title);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocumentWithContent($"test_cc_case_{operation}.docx", "Test");
        var outputPath = CreateTestFilePath($"test_cc_case_{operation}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, type: "PlainText",
            tag: $"cc_{operation}");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_cc_unknown.docx", "Test");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddContentControl_WithSessionId_ShouldWorkInMemory()
    {
        var docPath = CreateWordDocumentWithContent("test_cc_session_add.docx", "Test");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, type: "PlainText", tag: "sessionCC");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        Assert.True(sdtNodes.Count > 0);
    }

    [Fact]
    public void GetContentControls_WithSessionId_ShouldReturn()
    {
        var docPath = CreateWordDocument("test_cc_session_get.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline) { Tag = "sessionGet" };
        builder.InsertNode(sdt);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetContentControlsResult>(result);
        Assert.True(data.Count > 0);
        var output = GetResultOutput<GetContentControlsResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}
