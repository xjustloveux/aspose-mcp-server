using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Results.Word.Properties;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordPropertiesTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordPropertiesToolTests : WordTestBase
{
    private readonly WordPropertiesTool _tool;

    public WordPropertiesToolTests()
    {
        _tool = new WordPropertiesTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void GetProperties_ShouldReturnJsonFormat()
    {
        var docPath = CreateWordDocument("test_get_properties.docx");
        var result = _tool.Execute("get", docPath);
        Assert.NotNull(result);
        var data = GetResultData<GetWordPropertiesResult>(result);
        Assert.NotNull(data.BuiltInProperties);
        Assert.NotNull(data.Statistics);
    }

    [Fact]
    public void SetProperties_ShouldSetProperties()
    {
        var docPath = CreateWordDocument("test_set_properties.docx");
        var outputPath = CreateTestFilePath("test_set_properties_output.docx");
        _tool.Execute("set", docPath, outputPath: outputPath,
            title: "Test Document", author: "Test Author", subject: "Test Subject");
        var doc = new Document(outputPath);
        Assert.Equal("Test Document", doc.BuiltInDocumentProperties.Title);
        Assert.Equal("Test Author", doc.BuiltInDocumentProperties.Author);
        Assert.Equal("Test Subject", doc.BuiltInDocumentProperties.Subject);
    }

    [Fact]
    public void SetProperties_WithCustomStringProperty_ShouldAddAsString()
    {
        var docPath = CreateWordDocument("test_custom_string.docx");
        var outputPath = CreateTestFilePath("test_custom_string_output.docx");
        var customProps = new JsonObject
        {
            ["ProjectName"] = "My Project",
            ["Version"] = "1.0.0"
        };
        _tool.Execute("set", docPath, outputPath: outputPath, customProperties: customProps.ToJsonString());
        var doc = new Document(outputPath);
        Assert.Equal("My Project", doc.CustomDocumentProperties["ProjectName"]?.Value?.ToString());
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET")]
    [InlineData("GeT")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var result = _tool.Execute(operation, docPath);
        var data = GetResultData<GetWordPropertiesResult>(result);
        Assert.NotNull(data.BuiltInProperties);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_unknown_op.docx", "Test content");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetProperties_WithSessionId_ShouldReturnProperties()
    {
        var docPath = CreateWordDocument("test_session_get_props.docx");
        var doc = new Document(docPath) { BuiltInDocumentProperties = { Title = "Session Title" } };
        doc.Save(docPath);
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetWordPropertiesResult>(result);
        Assert.Equal("Session Title", data.BuiltInProperties.Title);
        var output = GetResultOutput<GetWordPropertiesResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void SetProperties_WithSessionId_ShouldSetPropertiesInMemory()
    {
        var docPath = CreateWordDocument("test_session_set_props.docx");
        var sessionId = OpenSession(docPath);
        _tool.Execute("set", sessionId: sessionId,
            title: "In-Memory Title", author: "In-Memory Author");
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("In-Memory Title", sessionDoc.BuiltInDocumentProperties.Title);
        Assert.Equal("In-Memory Author", sessionDoc.BuiltInDocumentProperties.Author);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_props.docx");
        var doc1 = new Document(docPath1) { BuiltInDocumentProperties = { Title = "Path Title" } };
        doc1.Save(docPath1);
        var docPath2 = CreateWordDocument("test_session_props.docx");
        var doc2 = new Document(docPath2) { BuiltInDocumentProperties = { Title = "Session Title" } };
        doc2.Save(docPath2);
        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);
        var data = GetResultData<GetWordPropertiesResult>(result);
        Assert.Equal("Session Title", data.BuiltInProperties.Title);
        Assert.NotEqual("Path Title", data.BuiltInProperties.Title);
    }

    #endregion
}
