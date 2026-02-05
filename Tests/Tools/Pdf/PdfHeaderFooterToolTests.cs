using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for <see cref="PdfHeaderFooterTool" />.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfHeaderFooterToolTests : PdfTestBase
{
    private readonly PdfHeaderFooterTool _tool;

    public PdfHeaderFooterToolTests()
    {
        _tool = new PdfHeaderFooterTool(SessionManager);
    }

    private string CreateTestPdf(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} Content"));
        }

        document.Save(filePath);
        return filePath;
    }

    private string CreateTestImage(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var width = 10;
        var height = 10;
        var bmp = new byte[width * height * 3 + 54];
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        var fileSize = bmp.Length;
        bmp[2] = (byte)(fileSize & 0xFF);
        bmp[3] = (byte)((fileSize >> 8) & 0xFF);
        bmp[4] = (byte)((fileSize >> 16) & 0xFF);
        bmp[5] = (byte)((fileSize >> 24) & 0xFF);
        bmp[10] = 54;
        bmp[14] = 40;
        bmp[18] = (byte)(width & 0xFF);
        bmp[19] = (byte)((width >> 8) & 0xFF);
        bmp[22] = (byte)(height & 0xFF);
        bmp[23] = (byte)((height >> 8) & 0xFF);
        bmp[26] = 1;
        bmp[28] = 24;
        for (var i = 54; i < bmp.Length; i += 3)
        {
            bmp[i] = 255;
            bmp[i + 1] = 0;
            bmp[i + 2] = 0;
        }

        File.WriteAllBytes(filePath, bmp);
        return filePath;
    }

    private string CreatePdfWithStampAnnotation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        var stamp = new StampAnnotation(page, new Rectangle(10, 10, 50, 50));
        page.Annotations.Add(stamp);
        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Execute_AddText_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_add_text.pdf");
        var outputPath = CreateTestFilePath("test_add_text_output.pdf");

        var result = _tool.Execute("add_text", pdfPath, outputPath: outputPath, text: "Header Text");

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("text", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_AddImage_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_add_image.pdf");
        var outputPath = CreateTestFilePath("test_add_image_output.pdf");
        var imagePath = CreateTestImage("test_header.bmp");

        var result = _tool.Execute("add_image", pdfPath, outputPath: outputPath, imagePath: imagePath);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("image", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_AddPageNumber_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_add_page_number.pdf", 3);
        var outputPath = CreateTestFilePath("test_add_page_number_output.pdf");

        var result = _tool.Execute("add_page_number", pdfPath, outputPath: outputPath);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("page numbers", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_Remove_ReturnsResult()
    {
        var pdfPath = CreatePdfWithStampAnnotation("test_remove.pdf");
        var outputPath = CreateTestFilePath("test_remove_output.pdf");

        var result = _tool.Execute("remove", pdfPath, outputPath: outputPath);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("stamp(s)", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_AddText_WithMultiplePages_AppliesAll()
    {
        var pdfPath = CreateTestPdf("test_add_text_multi.pdf", 3);
        var outputPath = CreateTestFilePath("test_add_text_multi_output.pdf");

        var result = _tool.Execute("add_text", pdfPath, outputPath: outputPath, text: "Multi-page Header");

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("3 page(s)", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_AddText_WithPageRange_AppliesOnlyToSelectedPages()
    {
        var pdfPath = CreateTestPdf("test_add_text_range.pdf", 4);
        var outputPath = CreateTestFilePath("test_add_text_range_output.pdf");

        var result = _tool.Execute("add_text", pdfPath, outputPath: outputPath,
            text: "Range Header", pageRange: "1-3");

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("3 page(s)", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD_TEXT")]
    [InlineData("Add_Text")]
    [InlineData("add_text")]
    public void AddText_Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");

        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath, text: "Test");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

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
        Assert.ThrowsAny<Exception>(() => _tool.Execute("add_text", text: "Test"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddText_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_add_text.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add_text", sessionId: sessionId, text: "Session Header");

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("text", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
    }

    [Fact]
    public void AddPageNumber_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_page_number.pdf", 3);
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("add_page_number", sessionId: sessionId);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("page numbers", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Remove_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreatePdfWithStampAnnotation("test_session_remove.pdf");
        var sessionId = OpenSession(pdfPath);

        var result = _tool.Execute("remove", sessionId: sessionId);

        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("stamp(s)", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("add_text", sessionId: "invalid_session", text: "Test"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_file.pdf");
        var pdfPath2 = CreateTestPdf("test_session_file.pdf", 3);
        var sessionId = OpenSession(pdfPath2);

        var result = _tool.Execute("add_text", pdfPath1, sessionId, text: "Test");

        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(3, document.Pages.Count);
    }

    #endregion
}
