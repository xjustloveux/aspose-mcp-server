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
    public void AddLink_ShouldAddLink()
    {
        var pdfPath = CreateTestPdf("test_add_link.pdf");
        var outputPath = CreateTestFilePath("test_add_link_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            x: 100,
            y: 100,
            width: 200,
            height: 30,
            url: "https://example.com");
        var document = new Document(outputPath);
        var page = document.Pages[1];
        var annotations = page.Annotations;
        Assert.True(annotations.Count > 0, "Page should contain at least one link annotation");
    }

    [Fact]
    public void GetLinks_ShouldReturnAllLinks()
    {
        var pdfPath = CreateTestPdf("test_get_links.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://test.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void DeleteLink_ShouldDeleteLink()
    {
        var pdfPath = CreateTestPdf("test_delete_link.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://delete.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var linksBefore = document.Pages[1].Annotations.Count;
        Assert.True(linksBefore > 0, "Link should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_link_output.pdf");
        _tool.Execute(
            "delete",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            linkIndex: 0);
        var resultDocument = new Document(outputPath);
        var linksAfter = resultDocument.Pages[1].Annotations.Count;
        Assert.True(linksAfter < linksBefore,
            $"Link should be deleted. Before: {linksBefore}, After: {linksAfter}");
    }

    [Fact]
    public void EditLink_ShouldEditLink()
    {
        var pdfPath = CreateTestPdf("test_edit_link.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://original.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_edit_link_output.pdf");
        _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            linkIndex: 0,
            url: "https://updated.com");
        Assert.True(File.Exists(outputPath), "Output file should be created");
        var resultDocument = new Document(outputPath);
        var annotations = resultDocument.Pages[1].Annotations;
        Assert.True(annotations.Count > 0, "Page should still have annotations");
    }

    [Fact]
    public void AddLink_WithTargetPage_ShouldAddInternalLink()
    {
        var pdfPath = CreateTestFilePath("test_add_internal_link.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_add_internal_link_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            x: 100,
            y: 100,
            width: 200,
            height: 30,
            targetPage: 2);
        var resultDocument = new Document(outputPath);
        var annotations = resultDocument.Pages[1].Annotations.OfType<LinkAnnotation>().ToList();
        Assert.True(annotations.Count > 0, "Page should contain internal link");
        Assert.IsType<GoToAction>(annotations[0].Action);
    }

    [Fact]
    public void GetLinks_WithInternalLink_ShouldReturnDestinationPage()
    {
        var pdfPath = CreateTestFilePath("test_get_internal_links.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToAction(document.Pages[2])
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        Assert.Contains("\"type\": \"page\"", result);
    }

    [Fact]
    public void GetLinks_WithNoLinks_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_no_links.pdf");
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No links found", result);
    }

    [Fact]
    public void GetLinks_WithoutPageIndex_ShouldReturnAllLinks()
    {
        var pdfPath = CreateTestFilePath("test_get_all_links.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        var page1 = document.Pages[1];
        var link1 = new LinkAnnotation(page1, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://page1.com")
        };
        page1.Annotations.Add(link1);

        var page2 = document.Pages[2];
        var link2 = new LinkAnnotation(page2, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://page2.com")
        };
        page2.Annotations.Add(link2);
        document.Save(pdfPath);
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 2", result);
    }

    [Fact]
    public void AddLink_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 99,
            x: 100,
            y: 100,
            width: 200,
            height: 30,
            url: "https://example.com"));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void AddLink_WithoutUrlOrTargetPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_url.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 30));
        Assert.Contains("Either url or targetPage must be provided", exception.Message);
    }

    [Fact]
    public void AddLink_WithInvalidTargetPage_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_target.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 30,
            targetPage: 99));
        Assert.Contains("targetPage must be between", exception.Message);
    }

    [Fact]
    public void DeleteLink_WithInvalidLinkIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_invalid_index.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://test.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            pageIndex: 1,
            linkIndex: 99));
        Assert.Contains("linkIndex must be between", exception.Message);
    }

    [Fact]
    public void EditLink_WithInvalidLinkIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_edit_invalid_index.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            linkIndex: 99,
            url: "https://test.com"));
        Assert.Contains("linkIndex must be between", exception.Message);
    }

    [Fact]
    public void EditLink_WithTargetPage_ShouldChangeToInternalLink()
    {
        var pdfPath = CreateTestFilePath("test_edit_to_internal.pdf");
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        var page = document.Pages[1];
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://original.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_edit_to_internal_output.pdf");
        _tool.Execute(
            "edit",
            pdfPath,
            pageIndex: 1,
            outputPath: outputPath,
            linkIndex: 0,
            targetPage: 2);
        var resultDocument = new Document(outputPath);
        var annotations = resultDocument.Pages[1].Annotations.OfType<LinkAnnotation>().ToList();
        Assert.True(annotations.Count > 0);
        Assert.IsType<GoToAction>(annotations[0].Action);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void GetLinks_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_get_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", pdfPath, pageIndex: 99));
        Assert.Contains("pageIndex must be between", exception.Message);
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
    public void AddLink_WithDefaultDimensions_ShouldCreateZeroSizeLink()
    {
        var pdfPath = CreateTestPdf("test_add_missing_dimensions.pdf");
        var outputPath = CreateTestFilePath("test_add_missing_dimensions_output.pdf");

        // Act - Default x, y, width, height are all 0, creating a point-size link area
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath, pageIndex: 1, url: "https://example.com");

        // Assert - Tool succeeds but creates a link with zero dimensions
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Added link to page", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetLinks_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestFilePath("test_session_get_links.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://session-test.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId, pageIndex: 1);
        Assert.NotNull(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void AddLink_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add_link.pdf");
        var sessionId = OpenSession(pdfPath);

        // Get initial annotation count
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var annotationCountBefore = docBefore.Pages[1].Annotations.Count;
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            pageIndex: 1,
            x: 100,
            y: 100,
            width: 200,
            height: 30,
            url: "https://session-add.com");
        Assert.Contains("Added link", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(docAfter);
        Assert.True(docAfter.Pages[1].Annotations.Count > annotationCountBefore,
            "Link should be added to session document");
    }

    [Fact]
    public void DeleteLink_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreateTestFilePath("test_session_delete_link.pdf");
        var document = new Document();
        var page = document.Pages.Add();
        var link = new LinkAnnotation(page, new Rectangle(100, 100, 300, 130))
        {
            Action = new GoToURIAction("https://session-delete.com")
        };
        page.Annotations.Add(link);
        document.Save(pdfPath);

        var sessionId = OpenSession(pdfPath);

        // Get initial annotation count
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var annotationCountBefore = docBefore.Pages[1].Annotations.Count;
        Assert.True(annotationCountBefore > 0, "Link should exist before deletion");
        var result = _tool.Execute(
            "delete",
            sessionId: sessionId,
            pageIndex: 1,
            linkIndex: 0);
        Assert.Contains("Deleted", result);

        // Verify in-memory changes
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docAfter.Pages[1].Annotations.Count < annotationCountBefore,
            "Link should be deleted from session document");
    }

    #endregion
}