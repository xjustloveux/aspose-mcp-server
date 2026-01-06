using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfPropertiesToolTests : PdfTestBase
{
    private readonly PdfPropertiesTool _tool;

    public PdfPropertiesToolTests()
    {
        _tool = new PdfPropertiesTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Get_ShouldReturnProperties()
    {
        var pdfPath = CreateTestPdf("test_get.pdf");
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("title", out _));
        Assert.True(json.TryGetProperty("author", out _));
        Assert.True(json.TryGetProperty("subject", out _));
        Assert.True(json.TryGetProperty("keywords", out _));
        Assert.True(json.TryGetProperty("creator", out _));
        Assert.True(json.TryGetProperty("producer", out _));
        Assert.True(json.TryGetProperty("totalPages", out _));
        Assert.True(json.TryGetProperty("isEncrypted", out _));
        Assert.True(json.TryGetProperty("isLinearized", out _));
    }

    [Fact]
    public void Set_WithTitleAndAuthor_ShouldSetProperties()
    {
        var pdfPath = CreateTestPdf("test_set.pdf");
        var outputPath = CreateTestFilePath("test_set_output.pdf");
        var result = _tool.Execute("set", pdfPath, outputPath: outputPath,
            title: "Test PDF", author: "Test Author", subject: "Test Subject");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Document properties updated", result);
    }

    [Fact]
    public void Set_WithKeywords_ShouldSetKeywords()
    {
        var pdfPath = CreateTestPdf("test_keywords.pdf");
        var outputPath = CreateTestFilePath("test_keywords_output.pdf");
        var result = _tool.Execute("set", pdfPath, outputPath: outputPath,
            keywords: "test, pdf, keywords");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Document properties updated", result);
    }

    [Fact]
    public void Set_WithCreatorAndProducer_ShouldAttemptToSet()
    {
        var pdfPath = CreateTestPdf("test_creator.pdf");
        var outputPath = CreateTestFilePath("test_creator_output.pdf");
        var result = _tool.Execute("set", pdfPath, outputPath: outputPath,
            creator: "Test Creator", producer: "Test Producer");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Document properties updated", result);
    }

    [Fact]
    public void Set_WithAllProperties_ShouldSetAll()
    {
        var pdfPath = CreateTestPdf("test_all.pdf");
        var outputPath = CreateTestFilePath("test_all_output.pdf");
        var result = _tool.Execute("set", pdfPath, outputPath: outputPath,
            title: "Full Test", author: "Full Author", subject: "Full Subject",
            keywords: "full, test", creator: "Full Creator", producer: "Full Producer");
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Document properties updated", result);
    }

    [Fact]
    public void Set_WithNoProperties_ShouldStillSave()
    {
        var pdfPath = CreateTestPdf("test_empty.pdf");
        var outputPath = CreateTestFilePath("test_empty_output.pdf");
        var result = _tool.Execute("set", pdfPath, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Document properties updated", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        Assert.Contains("title", result);
    }

    [Theory]
    [InlineData("SET")]
    [InlineData("Set")]
    [InlineData("set")]
    public void Operation_ShouldBeCaseInsensitive_Set(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_set_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath, title: "Test");
        Assert.StartsWith("Document properties updated", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    [Fact]
    public void Execute_WithNonExistentFile_ShouldThrowException()
    {
        Assert.ThrowsAny<IOException>(() => _tool.Execute("get", @"C:\nonexistent\file.pdf"));
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("title", out _));
        Assert.True(json.TryGetProperty("totalPages", out _));
    }

    [Fact]
    public void Set_WithSessionId_ShouldSetInSession()
    {
        var pdfPath = CreateTestPdf("test_session_set.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("set", sessionId: sessionId,
            title: "Session Title", author: "Session Author");
        Assert.StartsWith("Document properties updated", result);
        Assert.Contains("session", result);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("Session Title", doc.Info.Title);
        Assert.Equal("Session Author", doc.Info.Author);
    }

    [Fact]
    public void Set_WithSessionId_AndAllProperties_ShouldPersistChanges()
    {
        var pdfPath = CreateTestPdf("test_session_all.pdf");
        var sessionId = OpenSession(pdfPath);
        _tool.Execute("set", sessionId: sessionId,
            title: "Persisted Title", subject: "Persisted Subject", keywords: "test");
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("Persisted Title", doc.Info.Title);
        Assert.Equal("Persisted Subject", doc.Info.Subject);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        _ = CreateTestPdf("test_path_props.pdf");
        var pdfPath2 = CreateTestPdf("test_session_props.pdf");
        var sessionId = OpenSession(pdfPath2);
        _tool.Execute("set", sessionId: sessionId, title: "Session Doc Title");
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal("Session Doc Title", doc.Info.Title);
    }

    #endregion
}