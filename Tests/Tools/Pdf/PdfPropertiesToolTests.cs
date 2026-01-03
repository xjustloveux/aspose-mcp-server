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
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void GetProperties_ShouldReturnProperties()
    {
        var pdfPath = CreateTestPdf("test_get_properties.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("title", result);
    }

    [Fact]
    public void SetProperties_ShouldSetProperties()
    {
        var pdfPath = CreateTestPdf("test_set_properties.pdf");
        var outputPath = CreateTestFilePath("test_set_properties_output.pdf");
        _tool.Execute(
            "set",
            pdfPath,
            outputPath: outputPath,
            title: "Test PDF",
            author: "Test Author",
            subject: "Test Subject");
        using var document = new Document(outputPath);
        Assert.NotNull(document);
        Assert.True(document.Pages.Count > 0, "Document should have pages");
    }

    [Fact]
    public void GetProperties_ShouldReturnAllFields()
    {
        var pdfPath = CreateTestPdf("test_get_all_properties.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("title", result);
        Assert.Contains("author", result);
        Assert.Contains("subject", result);
        Assert.Contains("keywords", result);
        Assert.Contains("creator", result);
        Assert.Contains("producer", result);
        Assert.Contains("totalPages", result);
        Assert.Contains("isEncrypted", result);
        Assert.Contains("isLinearized", result);
    }

    [Fact]
    public void SetProperties_WithKeywords_ShouldSetKeywords()
    {
        var pdfPath = CreateTestPdf("test_set_keywords.pdf");
        var outputPath = CreateTestFilePath("test_set_keywords_output.pdf");
        var result = _tool.Execute(
            "set",
            pdfPath,
            outputPath: outputPath,
            keywords: "test, pdf, keywords");
        Assert.Contains("Document properties updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetProperties_WithCreatorAndProducer_ShouldAttemptToSet()
    {
        var pdfPath = CreateTestPdf("test_set_creator.pdf");
        var outputPath = CreateTestFilePath("test_set_creator_output.pdf");
        var result = _tool.Execute(
            "set",
            pdfPath,
            outputPath: outputPath,
            creator: "Test Creator",
            producer: "Test Producer");
        Assert.Contains("Document properties updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetProperties_WithAllProperties_ShouldSetAll()
    {
        var pdfPath = CreateTestPdf("test_set_all.pdf");
        var outputPath = CreateTestFilePath("test_set_all_output.pdf");
        var result = _tool.Execute(
            "set",
            pdfPath,
            outputPath: outputPath,
            title: "Full Test",
            author: "Full Author",
            subject: "Full Subject",
            keywords: "full, test",
            creator: "Full Creator",
            producer: "Full Producer");
        Assert.Contains("Document properties updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetProperties_WithNoProperties_ShouldStillSave()
    {
        var pdfPath = CreateTestPdf("test_set_empty.pdf");
        var outputPath = CreateTestFilePath("test_set_empty_output.pdf");
        var result = _tool.Execute(
            "set",
            pdfPath,
            outputPath: outputPath);
        Assert.Contains("Document properties updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithMissingRequiredPath_ShouldThrowArgumentException()
    {
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("get"));
        Assert.Contains("path", exception.Message.ToLower());
    }

    [Fact]
    public void Execute_WithNonExistentFile_ShouldThrowException()
    {
        // Act & Assert - May throw FileNotFoundException or DirectoryNotFoundException depending on path
        Assert.ThrowsAny<IOException>(() => _tool.Execute("get", @"C:\nonexistent\file.pdf"));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetProperties_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get_properties.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("title", result);
        Assert.Contains("totalPages", result);
    }

    [Fact]
    public void SetProperties_WithSessionId_ShouldSetInSession()
    {
        var pdfPath = CreateTestPdf("test_session_set_properties.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "set",
            sessionId: sessionId,
            title: "Session Title",
            author: "Session Author");
        Assert.Contains("Document properties updated", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc);
        Assert.Equal("Session Title", doc.Info.Title);
        Assert.Equal("Session Author", doc.Info.Author);
    }

    [Fact]
    public void SetProperties_WithSessionId_ShouldPersistChanges()
    {
        var pdfPath = CreateTestPdf("test_session_persist_properties.pdf");
        var sessionId = OpenSession(pdfPath);

        // Set properties
        _tool.Execute(
            "set",
            sessionId: sessionId,
            title: "Persisted Title",
            subject: "Persisted Subject");

        // Assert - Verify in-memory document has the properties set
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc);
        Assert.Equal("Persisted Title", doc.Info.Title);
        Assert.Equal("Persisted Subject", doc.Info.Subject);
    }

    #endregion
}