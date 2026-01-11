using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfPageTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfPageToolTests : PdfTestBase
{
    private readonly PdfPageTool _tool;

    public PdfPageToolTests()
    {
        _tool = new PdfPageTool(SessionManager);
    }

    private string CreateTestPdf(string fileName, int pageCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddPageAndPersistToFile()
    {
        var pdfPath = CreateTestPdf("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath, count: 1);

        Assert.StartsWith("Added 1 page(s)", result);
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);
    }

    [Fact]
    public void Delete_ShouldDeletePageAndPersistToFile()
    {
        var pdfPath = CreateTestPdf("test_delete.pdf");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");

        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath, pageIndex: 1);

        Assert.StartsWith("Deleted page 1", result);
        using var document = new Document(outputPath);
        Assert.Single(document.Pages);
    }

    [Fact]
    public void Rotate_ShouldRotatePageAndPersistToFile()
    {
        var pdfPath = CreateTestPdf("test_rotate.pdf");
        var outputPath = CreateTestFilePath("test_rotate_output.pdf");

        var result = _tool.Execute("rotate", pdfPath, outputPath: outputPath,
            pageIndex: 1, rotation: 90);

        Assert.StartsWith("Rotated 1 page(s) by 90 degrees", result);
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.on90, document.Pages[1].Rotate);
    }

    [Fact]
    public void GetInfo_ShouldReturnPageInfoFromFile()
    {
        var pdfPath = CreateTestPdf("test_info.pdf");
        var result = _tool.Execute("get_info", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(2, json.GetProperty("count").GetInt32());
        Assert.True(json.TryGetProperty("items", out var items));
        Assert.Equal(2, items.GetArrayLength());
    }

    [Fact]
    public void GetDetails_ShouldReturnPageDetailsFromFile()
    {
        var pdfPath = CreateTestPdf("test_details.pdf");
        var result = _tool.Execute("get_details", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(1, json.GetProperty("pageIndex").GetInt32());
        Assert.True(json.TryGetProperty("width", out _));
        Assert.True(json.TryGetProperty("height", out _));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");

        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath, count: 1);

        Assert.StartsWith("Added 1 page(s)", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_info"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetInfo_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_info.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("get_info", sessionId: sessionId);
        var json = JsonSerializer.Deserialize<JsonElement>(result);

        Assert.Equal(2, json.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var countBefore = docBefore.Pages.Count;

        var result = _tool.Execute("add", sessionId: sessionId, count: 1);

        Assert.StartsWith("Added 1 page(s)", result);
        Assert.Contains(sessionId, result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(countBefore + 1, docAfter.Pages.Count);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_delete.pdf");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var countBefore = docBefore.Pages.Count;

        var result = _tool.Execute("delete", sessionId: sessionId, pageIndex: 1);

        Assert.StartsWith("Deleted page 1", result);
        Assert.Contains(sessionId, result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(countBefore - 1, docAfter.Pages.Count);
    }

    [Fact]
    public void Rotate_WithSessionId_ShouldRotateInSession()
    {
        var pdfPath = CreateTestPdf("test_session_rotate.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("rotate", sessionId: sessionId, pageIndex: 1, rotation: 90);

        Assert.StartsWith("Rotated 1 page(s) by 90 degrees", result);
        Assert.Contains(sessionId, result);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(Rotation.on90, doc.Pages[1].Rotate);
    }

    [Fact]
    public void GetDetails_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_details.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("get_details", sessionId: sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);

        Assert.Equal(1, json.GetProperty("pageIndex").GetInt32());
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_info", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_page.pdf", 1);
        var pdfPath2 = CreateTestPdf("test_session_page.pdf", 3);
        var sessionId = OpenSession(pdfPath2);

        var result = _tool.Execute("get_info", pdfPath1, sessionId);
        var json = JsonSerializer.Deserialize<JsonElement>(result);

        Assert.Equal(3, json.GetProperty("count").GetInt32());
    }

    #endregion
}
