using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfLinkToolTests : PdfTestBase
{
    private readonly PdfLinkTool _tool;

    public PdfLinkToolTests()
    {
        _tool = new PdfLinkTool(SessionManager);
    }

    private string CreateTestPdf(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreatePdfWithLink(string fileName, string url = "https://test.com")
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction(url)
        };
        page.Annotations.Add(link);
        document.Save(filePath);
        return filePath;
    }

    private string CreatePdfWithInternalLink(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToAction(document.Pages[2])
        };
        page.Annotations.Add(link);
        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Add_WithUrl_ShouldAddExternalLink()
    {
        var pdfPath = CreateTestPdf("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 30, url: "https://example.com");
        Assert.StartsWith("Added link to page 1", result);
        using var document = new Document(outputPath);
        Assert.True(document.Pages[1].Annotations.Count > 0);
    }

    [Fact]
    public void Add_WithTargetPage_ShouldAddInternalLink()
    {
        var pdfPath = CreateTestPdf("test_add_internal.pdf", 2);
        var outputPath = CreateTestFilePath("test_add_internal_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 30, targetPage: 2);
        Assert.StartsWith("Added link to page 1", result);
        using var document = new Document(outputPath);
        var links = document.Pages[1].Annotations.OfType<LinkAnnotation>().ToList();
        Assert.NotEmpty(links);
        Assert.IsType<GoToAction>(links[0].Action);
    }

    [Fact]
    public void Delete_ShouldDeleteLink()
    {
        var pdfPath = CreatePdfWithLink("test_delete.pdf");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");
        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath,
            pageIndex: 1, linkIndex: 0);
        Assert.StartsWith("Deleted link 0", result);
        using var document = new Document(outputPath);
        Assert.Empty(document.Pages[1].Annotations.OfType<LinkAnnotation>());
    }

    [Fact]
    public void Edit_WithUrl_ShouldUpdateLink()
    {
        var pdfPath = CreatePdfWithLink("test_edit.pdf", "https://original.com");
        var outputPath = CreateTestFilePath("test_edit_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            pageIndex: 1, linkIndex: 0, url: "https://updated.com");
        Assert.StartsWith("Edited link 0", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Edit_WithTargetPage_ShouldChangeToInternalLink()
    {
        var pdfPath = CreateTestFilePath("test_edit_internal.pdf");
        using (var document = new Document())
        {
            document.Pages.Add();
            document.Pages.Add();
            var page = document.Pages[1];
            var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
            {
                Action = new GoToURIAction("https://original.com")
            };
            page.Annotations.Add(link);
            document.Save(pdfPath);
        }

        var outputPath = CreateTestFilePath("test_edit_internal_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            pageIndex: 1, linkIndex: 0, targetPage: 2);
        Assert.StartsWith("Edited link 0", result);
        using var resultDoc = new Document(outputPath);
        var links = resultDoc.Pages[1].Annotations.OfType<LinkAnnotation>().ToList();
        Assert.IsType<GoToAction>(links[0].Action);
    }

    [Fact]
    public void Get_WithLinks_ShouldReturnLinks()
    {
        var pdfPath = CreatePdfWithLink("test_get.pdf");
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
        Assert.True(json.TryGetProperty("items", out _));
    }

    [Fact]
    public void Get_WithNoLinks_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(0, json.GetProperty("count").GetInt32());
        Assert.Contains("No links found", result);
    }

    [Fact]
    public void Get_WithInternalLink_ShouldReturnDestinationPage()
    {
        var pdfPath = CreatePdfWithInternalLink("test_get_internal.pdf");
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        Assert.Contains("\"type\": \"page\"", result);
    }

    [Fact]
    public void Get_WithoutPageIndex_ShouldReturnAllLinks()
    {
        var pdfPath = CreateTestFilePath("test_get_all.pdf");
        using (var document = new Document())
        {
            document.Pages.Add();
            document.Pages.Add();
            document.Pages[1].Annotations.Add(new LinkAnnotation(document.Pages[1], new Rectangle(100, 100, 300, 130))
            {
                Action = new GoToURIAction("https://page1.com")
            });
            document.Pages[2].Annotations.Add(new LinkAnnotation(document.Pages[2], new Rectangle(100, 100, 300, 130))
            {
                Action = new GoToURIAction("https://page2.com")
            });
            document.Save(pdfPath);
        }

        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(2, json.GetProperty("count").GetInt32());
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 30, url: "https://example.com");
        Assert.StartsWith("Added link", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_get_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath, pageIndex: 1);
        Assert.Contains("count", result);
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
    public void Add_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 99, x: 100, y: 100, width: 200, height: 30,
                url: "https://test.com"));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Add_WithoutUrlOrTargetPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_url.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1, x: 100, y: 100, width: 200, height: 30));
        Assert.Contains("Either url or targetPage must be provided", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidTargetPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_target.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, pageIndex: 1, x: 100, y: 100, width: 200, height: 30, targetPage: 99));
        Assert.Contains("targetPage must be between", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidLinkIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithLink("test_delete_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, pageIndex: 1, linkIndex: 99));
        Assert.Contains("linkIndex must be between", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidLinkIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_edit_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, pageIndex: 1, linkIndex: 99, url: "https://test.com"));
        Assert.Contains("linkIndex must be between", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_get_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", pdfPath, pageIndex: 99));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreatePdfWithLink("test_session_get.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var countBefore = docBefore.Pages[1].Annotations.Count;
        var result = _tool.Execute("add", sessionId: sessionId,
            pageIndex: 1, x: 100, y: 100, width: 200, height: 30, url: "https://session.com");
        Assert.StartsWith("Added link", result);
        Assert.Contains("session", result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docAfter.Pages[1].Annotations.Count > countBefore);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithLink("test_session_delete.pdf");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var countBefore = docBefore.Pages[1].Annotations.Count;
        var result = _tool.Execute("delete", sessionId: sessionId, pageIndex: 1, linkIndex: 0);
        Assert.StartsWith("Deleted", result);
        Assert.Contains("session", result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docAfter.Pages[1].Annotations.Count < countBefore);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInSession()
    {
        var pdfPath = CreatePdfWithLink("test_session_edit.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("edit", sessionId: sessionId,
            pageIndex: 1, linkIndex: 0, url: "https://updated.com");
        Assert.StartsWith("Edited link", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", pageIndex: 1));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_link.pdf");
        var pdfPath2 = CreatePdfWithLink("test_session_link.pdf");
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get", pdfPath1, sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
    }

    #endregion
}