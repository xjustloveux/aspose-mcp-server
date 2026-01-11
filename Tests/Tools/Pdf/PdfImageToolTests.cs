using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;
using DrawingColor = System.Drawing.Color;
using Rectangle = Aspose.Pdf.Rectangle;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfImageTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
[SupportedOSPlatform("windows")]
public class PdfImageToolTests : PdfTestBase
{
    private readonly PdfImageTool _tool;

    public PdfImageToolTests()
    {
        _tool = new PdfImageTool(SessionManager);
    }

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(100, 100);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(DrawingColor.Red);
        }

        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
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

    private string CreatePdfWithImage(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        var imagePath = CreateTestImage($"img_{Guid.NewGuid()}.png");
        using (var imageStream = new FileStream(imagePath, FileMode.Open))
        {
            page.AddImage(imageStream, new Rectangle(100, 100, 300, 300));
        }

        document.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddImageToPdf()
    {
        var pdfPath = CreateTestPdf("test_add.pdf");
        var imagePath = CreateTestImage("test_add_image.png");
        var outputPath = CreateTestFilePath("test_add_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            pageIndex: 1, imagePath: imagePath, x: 100, y: 100);
        Assert.StartsWith("Added image", result);
        using var document = new Document(outputPath);
        Assert.True(document.Pages[1].Resources.Images.Count > 0);
    }

    [Fact]
    public void Delete_ShouldDeleteImage()
    {
        var pdfPath = CreatePdfWithImage("test_delete.pdf");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");
        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath,
            pageIndex: 1, imageIndex: 1);
        Assert.StartsWith("Deleted image", result);
        using var document = new Document(outputPath);
        Assert.Empty(document.Pages[1].Resources.Images);
    }

    [Fact]
    public void Extract_ShouldExtractImage()
    {
        var pdfPath = CreatePdfWithImage("test_extract.pdf");
        var outputImagePath = CreateTestFilePath("test_extracted.png");
        var result = _tool.Execute("extract", pdfPath,
            outputPath: outputImagePath, pageIndex: 1, imageIndex: 1);
        Assert.StartsWith("Extracted image", result);
        Assert.True(File.Exists(outputImagePath));
    }

    [Fact]
    public void Get_WithImages_ShouldReturnImageInfo()
    {
        var pdfPath = CreatePdfWithImage("test_get.pdf");
        var result = _tool.Execute("get", pdfPath, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
        Assert.True(json.TryGetProperty("items", out _));
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
        var imagePath = CreateTestImage($"test_case_{operation}.png");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            pageIndex: 1, imagePath: imagePath, x: 100, y: 100);
        Assert.StartsWith("Added", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreatePdfWithImage("test_session_get.pdf");
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
        var imagePath = CreateTestImage("test_session_add.png");
        var result = _tool.Execute("add", sessionId: sessionId,
            pageIndex: 1, imagePath: imagePath, x: 100, y: 100);
        Assert.StartsWith("Added image", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithImage("test_session_delete.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("delete", sessionId: sessionId,
            pageIndex: 1, imageIndex: 1);
        Assert.StartsWith("Deleted", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session", pageIndex: 1));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_image.pdf");
        var pdfPath2 = CreatePdfWithImage("test_session_image.pdf");
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get", pdfPath1, sessionId, pageIndex: 1);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
    }

    #endregion
}
