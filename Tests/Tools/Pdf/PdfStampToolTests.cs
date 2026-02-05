using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Pdf.Stamp;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfStampTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PdfStampToolTests : PdfTestBase
{
    private readonly PdfStampTool _tool;

    public PdfStampToolTests()
    {
        _tool = new PdfStampTool(SessionManager);
    }

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreateStampSourcePdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Stamp Source"));
        document.Save(filePath);
        return filePath;
    }

    #region Helper Methods

    private string CreateTempBmpImage()
    {
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

        var imagePath = Path.Combine(TestDir, $"temp_{Guid.NewGuid()}.bmp");
        File.WriteAllBytes(imagePath, bmp);
        return imagePath;
    }

    #endregion

    #region File I/O Smoke Tests

    [Fact]
    public void Execute_AddText_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_add_text.pdf");
        var outputPath = CreateTestFilePath("test_add_text_output.pdf");
        var result = _tool.Execute("add_text", pdfPath, outputPath: outputPath,
            text: "CONFIDENTIAL", pageIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Text stamp added", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_AddImage_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_add_image.pdf");
        var outputPath = CreateTestFilePath("test_add_image_output.pdf");
        var imagePath = CreateTempBmpImage();
        var result = _tool.Execute("add_image", pdfPath, outputPath: outputPath,
            imagePath: imagePath, pageIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Image stamp added", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_AddPdf_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_add_pdf.pdf");
        var stampSourcePath = CreateStampSourcePdf("stamp_source.pdf");
        var outputPath = CreateTestFilePath("test_add_pdf_output.pdf");
        var result = _tool.Execute("add_pdf", pdfPath, outputPath: outputPath,
            pdfPath: stampSourcePath, pageIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("PDF page stamp added", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_List_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_list.pdf");
        var result = _tool.Execute("list", pdfPath);
        var data = GetResultData<GetStampsPdfResult>(result);
        Assert.NotNull(data.Stamps);
        Assert.Equal(0, data.Count);
    }

    [Fact]
    public void Execute_Remove_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_remove.pdf");
        var outputPath = CreateTestFilePath("test_remove_output.pdf");
        var result = _tool.Execute("remove", pdfPath, outputPath: outputPath, pageIndex: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("No stamp annotations found", data.Message);
    }

    [Fact]
    public void Execute_AddText_WithAllOptions_ReturnsResult()
    {
        var pdfPath = CreateTestPdf("test_add_text_all.pdf");
        var outputPath = CreateTestFilePath("test_add_text_all_output.pdf");
        var result = _tool.Execute("add_text", pdfPath, outputPath: outputPath,
            text: "DRAFT", pageIndex: 1, x: 100, y: 200,
            fontSize: 20.0, opacity: 0.5, rotation: 45.0, color: "red");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Text stamp added", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD_TEXT")]
    [InlineData("Add_Text")]
    [InlineData("add_text")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath, text: "Test");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Text stamp added", data.Message);
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
        Assert.ThrowsAny<Exception>(() => _tool.Execute("list"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void List_WithSessionId_ShouldListFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_list.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("list", sessionId: sessionId);
        var data = GetResultData<GetStampsPdfResult>(result);
        Assert.NotNull(data.Stamps);
    }

    [Fact]
    public void AddText_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add_text.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("add_text", sessionId: sessionId,
            text: "Session Stamp", pageIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Text stamp added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("list", sessionId: "invalid_session"));
    }

    #endregion
}
