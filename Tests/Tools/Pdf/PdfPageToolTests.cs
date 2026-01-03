using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfPageToolTests : PdfTestBase
{
    private readonly PdfPageTool _tool;

    public PdfPageToolTests()
    {
        _tool = new PdfPageTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddPage_ShouldAddPage()
    {
        var pdfPath = CreateTestPdf("test_add_page.pdf");
        var pagesBefore = new Document(pdfPath).Pages.Count;
        var outputPath = CreateTestFilePath("test_add_page_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            count: 1);
        using var document = new Document(outputPath);
        Assert.True(document.Pages.Count > pagesBefore, "Page should be added");
    }

    [Fact]
    public void DeletePage_ShouldDeletePage()
    {
        var pdfPath = CreateTestPdf("test_delete_page.pdf");
        var pagesBefore = new Document(pdfPath).Pages.Count;
        Assert.True(pagesBefore >= 2, "PDF should have at least 2 pages");

        var outputPath = CreateTestFilePath("test_delete_page_output.pdf");
        _tool.Execute(
            "delete",
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1);
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count < pagesBefore, "Page should be deleted");
    }

    [Fact]
    public void RotatePage_ShouldRotatePage()
    {
        var pdfPath = CreateTestPdf("test_rotate_page.pdf");
        var outputPath = CreateTestFilePath("test_rotate_page_output.pdf");
        _tool.Execute(
            "rotate",
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            rotation: 90);
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public void GetPageInfo_ShouldReturnPageInfo()
    {
        var pdfPath = CreateTestPdf("test_get_page_info.pdf");
        var result = _tool.Execute("get_info", pdfPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void GetPageDetails_ShouldReturnPageDetails()
    {
        var pdfPath = CreateTestPdf("test_get_page_details.pdf");
        var result = _tool.Execute("get_details", pdfPath, pageIndex: 1);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("pageIndex", result);
        Assert.Contains("width", result);
        Assert.Contains("height", result);
        Assert.Contains("mediaBox", result);
        Assert.Contains("cropBox", result);
    }

    [Fact]
    public void GetPageInfo_ShouldReturnAllPagesInfo()
    {
        var pdfPath = CreateTestPdf("test_get_all_pages_info.pdf");
        var result = _tool.Execute("get_info", pdfPath);
        Assert.NotNull(result);
        Assert.Contains("\"count\": 2", result);
        Assert.Contains("items", result);
    }

    [Fact]
    public void AddPage_WithCustomSize_ShouldAddPageWithSize()
    {
        var pdfPath = CreateTestPdf("test_add_custom_size.pdf");
        var outputPath = CreateTestFilePath("test_add_custom_size_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            count: 1,
            width: 400,
            height: 600);
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);
    }

    [Fact]
    public void AddPage_WithInsertAt_ShouldInsertAtPosition()
    {
        var pdfPath = CreateTestPdf("test_add_insert_at.pdf");
        var outputPath = CreateTestFilePath("test_add_insert_at_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            count: 1,
            insertAt: 1);
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);
    }

    [SkippableFact]
    public void AddPage_WithMultiplePages_ShouldAddMultiplePages()
    {
        // Skip in evaluation mode - adding 3 pages to 2-page PDF exceeds 4-page limit
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Adding 3 pages to 2-page PDF exceeds 4-page limit");
        var pdfPath = CreateTestPdf("test_add_multiple.pdf");
        var outputPath = CreateTestFilePath("test_add_multiple_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            count: 3);
        using var document = new Document(outputPath);
        Assert.Equal(5, document.Pages.Count);
    }

    [Fact]
    public void DeletePage_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_invalid.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            pageIndex: 99));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void RotatePage_WithInvalidRotation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_rotate_invalid.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "rotate",
            pdfPath,
            pageIndex: 1,
            rotation: 45));
        Assert.Contains("rotation must be 0, 90, 180, or 270", exception.Message);
    }

    [Fact]
    public void RotatePage_WithPageIndices_ShouldRotateMultiplePages()
    {
        var pdfPath = CreateTestPdf("test_rotate_multiple.pdf");
        var outputPath = CreateTestFilePath("test_rotate_multiple_output.pdf");
        var result = _tool.Execute(
            "rotate",
            pdfPath,
            outputPath: outputPath,
            rotation: 90,
            pageIndices: [1, 2]);
        Assert.Contains("Rotated 2 page(s)", result);
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.on90, document.Pages[1].Rotate);
        Assert.Equal(Rotation.on90, document.Pages[2].Rotate);
    }

    [Fact]
    public void RotatePage_WithoutPageIndex_ShouldRotateAllPages()
    {
        var pdfPath = CreateTestPdf("test_rotate_all.pdf");
        var outputPath = CreateTestFilePath("test_rotate_all_output.pdf");
        var result = _tool.Execute(
            "rotate",
            pdfPath,
            outputPath: outputPath,
            rotation: 180);
        Assert.Contains("Rotated 2 page(s)", result);
    }

    [Fact]
    public void RotatePage_With270Degrees_ShouldRotate270()
    {
        var pdfPath = CreateTestPdf("test_rotate_270.pdf");
        var outputPath = CreateTestFilePath("test_rotate_270_output.pdf");
        _tool.Execute(
            "rotate",
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            rotation: 270);
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.on270, document.Pages[1].Rotate);
    }

    [Fact]
    public void RotatePage_With0Degrees_ShouldResetRotation()
    {
        var pdfPath = CreateTestPdf("test_rotate_0.pdf");
        var outputPath = CreateTestFilePath("test_rotate_0_output.pdf");
        _tool.Execute(
            "rotate",
            pdfPath,
            outputPath: outputPath,
            pageIndex: 1,
            rotation: 0);
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.None, document.Pages[1].Rotate);
    }

    [Fact]
    public void GetPageDetails_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_details_invalid.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "get_details",
            pdfPath,
            pageIndex: 99));
        Assert.Contains("pageIndex must be between", exception.Message);
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
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("get_info"));
        Assert.Contains("path", exception.Message.ToLower());
    }

    [Fact]
    public void RotatePage_WithDefaultRotation_ShouldRotateToZero()
    {
        var pdfPath = CreateTestPdf("test_rotate_missing.pdf");
        var outputPath = CreateTestFilePath("test_rotate_missing_output.pdf");

        // Act - Default rotation is 0 degrees, which is valid
        var result = _tool.Execute("rotate", pdfPath, outputPath: outputPath, pageIndex: 1);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Rotated 1 page(s)", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetPageInfo_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get_info.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get_info", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public void AddPage_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add_page.pdf");
        var sessionId = OpenSession(pdfPath);

        // Get initial page count
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var pageCountBefore = docBefore.Pages.Count;
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            count: 1);
        Assert.Contains("Added", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(docAfter);
        Assert.Equal(pageCountBefore + 1, docAfter.Pages.Count);
    }

    [Fact]
    public void DeletePage_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_delete_page.pdf");
        var sessionId = OpenSession(pdfPath);

        // Get initial page count
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var pageCountBefore = docBefore.Pages.Count;
        Assert.True(pageCountBefore >= 2, "PDF should have at least 2 pages");
        var result = _tool.Execute(
            "delete",
            sessionId: sessionId,
            pageIndex: 1);
        Assert.Contains("Deleted", result);

        // Verify in-memory changes
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(pageCountBefore - 1, docAfter.Pages.Count);
    }

    [Fact]
    public void RotatePage_WithSessionId_ShouldRotateInSession()
    {
        var pdfPath = CreateTestPdf("test_session_rotate_page.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "rotate",
            sessionId: sessionId,
            pageIndex: 1,
            rotation: 90);
        Assert.Contains("Rotated", result);

        // Verify in-memory changes
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc);
        Assert.Equal(Rotation.on90, doc.Pages[1].Rotate);
    }

    #endregion
}