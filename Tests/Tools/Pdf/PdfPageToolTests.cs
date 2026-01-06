using System.Text.Json;
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

    private string CreateTestPdf(string fileName, int pageCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddPage()
    {
        var pdfPath = CreateTestPdf("test_add.pdf");
        var outputPath = CreateTestFilePath("test_add_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath, count: 1);

        Assert.StartsWith("Added 1 page(s)", result);
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);
    }

    [Fact]
    public void Add_WithCustomSize_ShouldAddPageWithSize()
    {
        var pdfPath = CreateTestPdf("test_add_size.pdf");
        var outputPath = CreateTestFilePath("test_add_size_output.pdf");
        const double expectedWidth = 400;
        const double expectedHeight = 600;

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            count: 1, width: expectedWidth, height: expectedHeight);

        Assert.StartsWith("Added 1 page(s)", result);
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);

        var newPage = document.Pages[document.Pages.Count];
        var pageWidth = newPage.MediaBox.Width;
        var pageHeight = newPage.MediaBox.Height;

        Assert.True(Math.Abs(pageWidth - expectedWidth) < 1.0,
            $"Page width {pageWidth} should be approximately {expectedWidth}");
        Assert.True(Math.Abs(pageHeight - expectedHeight) < 1.0,
            $"Page height {pageHeight} should be approximately {expectedHeight}");
    }

    [Fact]
    public void Add_WithInsertAt_ShouldInsertAtPosition()
    {
        var pdfPath = CreateTestPdf("test_insert.pdf");
        var outputPath = CreateTestFilePath("test_insert_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            count: 1, insertAt: 1);

        Assert.StartsWith("Added 1 page(s)", result);
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);
    }

    [SkippableFact]
    public void Add_WithMultiplePages_ShouldAddMultiplePages()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Adding 3 pages to 2-page PDF exceeds 4-page limit");
        var pdfPath = CreateTestPdf("test_add_multi.pdf");
        var outputPath = CreateTestFilePath("test_add_multi_output.pdf");

        var result = _tool.Execute("add", pdfPath, outputPath: outputPath, count: 3);

        Assert.StartsWith("Added 3 page(s)", result);
        using var document = new Document(outputPath);
        Assert.Equal(5, document.Pages.Count);
    }

    [Fact]
    public void Delete_ShouldDeletePage()
    {
        var pdfPath = CreateTestPdf("test_delete.pdf");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");

        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath, pageIndex: 1);

        Assert.StartsWith("Deleted page 1", result);
        using var document = new Document(outputPath);
        Assert.Single(document.Pages);
    }

    [Fact]
    public void Rotate_ShouldRotatePage()
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
    public void Rotate_WithPageIndices_ShouldRotateMultiplePages()
    {
        var pdfPath = CreateTestPdf("test_rotate_multi.pdf");
        var outputPath = CreateTestFilePath("test_rotate_multi_output.pdf");

        var result = _tool.Execute("rotate", pdfPath, outputPath: outputPath,
            rotation: 90, pageIndices: [1, 2]);

        Assert.StartsWith("Rotated 2 page(s) by 90 degrees", result);
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.on90, document.Pages[1].Rotate);
        Assert.Equal(Rotation.on90, document.Pages[2].Rotate);
    }

    [Fact]
    public void Rotate_WithoutPageIndex_ShouldRotateAllPages()
    {
        var pdfPath = CreateTestPdf("test_rotate_all.pdf");
        var outputPath = CreateTestFilePath("test_rotate_all_output.pdf");

        var result = _tool.Execute("rotate", pdfPath, outputPath: outputPath, rotation: 180);

        Assert.StartsWith("Rotated 2 page(s) by 180 degrees", result);
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.on180, document.Pages[1].Rotate);
        Assert.Equal(Rotation.on180, document.Pages[2].Rotate);
    }

    [Theory]
    [InlineData(0, Rotation.None)]
    [InlineData(90, Rotation.on90)]
    [InlineData(180, Rotation.on180)]
    [InlineData(270, Rotation.on270)]
    public void Rotate_WithValidAngles_ShouldApplyCorrectRotation(int angle, Rotation expected)
    {
        var pdfPath = CreateTestPdf($"test_rotate_{angle}.pdf");
        var outputPath = CreateTestFilePath($"test_rotate_{angle}_output.pdf");
        _tool.Execute("rotate", pdfPath, outputPath: outputPath, pageIndex: 1, rotation: angle);
        using var document = new Document(outputPath);
        Assert.Equal(expected, document.Pages[1].Rotate);
    }

    [Fact]
    public void GetInfo_ShouldReturnPageInfo()
    {
        var pdfPath = CreateTestPdf("test_info.pdf");
        var result = _tool.Execute("get_info", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(2, json.GetProperty("count").GetInt32());
        Assert.True(json.TryGetProperty("items", out var items));
        Assert.Equal(2, items.GetArrayLength());
    }

    [Fact]
    public void GetDetails_ShouldReturnPageDetails()
    {
        var pdfPath = CreateTestPdf("test_details.pdf");
        var result = _tool.Execute("get_details", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(1, json.GetProperty("pageIndex").GetInt32());
        Assert.True(json.TryGetProperty("width", out _));
        Assert.True(json.TryGetProperty("height", out _));
        Assert.True(json.TryGetProperty("mediaBox", out _));
        Assert.True(json.TryGetProperty("cropBox", out _));
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");

        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath, count: 1);

        Assert.StartsWith("Added 1 page(s)", result);
    }

    [Theory]
    [InlineData("GET_INFO")]
    [InlineData("Get_Info")]
    [InlineData("get_info")]
    public void Operation_ShouldBeCaseInsensitive_GetInfo(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_info_{operation.Replace("_", "")}.pdf");

        var result = _tool.Execute(operation, pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);

        Assert.True(json.TryGetProperty("count", out var countProp));
        Assert.Equal(2, countProp.GetInt32());
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.StartsWith("Unknown operation: unknown", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, pageIndex: 99));
        Assert.StartsWith("pageIndex must be between 1 and", ex.Message);
    }

    [Fact]
    public void Rotate_WithInvalidRotation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_rotate_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("rotate", pdfPath, pageIndex: 1, rotation: 45));
        Assert.Equal("rotation must be 0, 90, 180, or 270", ex.Message);
    }

    [Fact]
    public void GetDetails_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_details_invalid.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_details", pdfPath, pageIndex: 99));
        Assert.StartsWith("pageIndex must be between 1 and", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_info"));
    }

    #endregion

    #region Session

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