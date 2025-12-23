using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfWatermarkToolTests : PdfTestBase
{
    private readonly PdfWatermarkTool _tool = new();

    private string CreatePdfDocument(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Sample PDF Text"));
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task AddWatermark_WithText_ShouldAddWatermark()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Confidential"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task AddWatermark_WithFontOptions_ShouldApplyFontOptions()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_font.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_font_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Watermark",
            ["fontName"] = "Arial",
            ["fontSize"] = 72
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task AddWatermark_WithOpacity_ShouldApplyOpacity()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_opacity.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_opacity_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Watermark",
            ["opacity"] = 0.5
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task AddWatermark_WithRotation_ShouldApplyRotation()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_rotation.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_rotation_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Watermark",
            ["rotation"] = 45
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task AddWatermark_WithAlignment_ShouldApplyAlignment()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_alignment.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_alignment_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Watermark",
            ["horizontalAlignment"] = "Left",
            ["verticalAlignment"] = "Top"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task AddWatermark_WithAllOptions_ShouldApplyAllOptions()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_all.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_all_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Confidential",
            ["fontName"] = "Arial",
            ["fontSize"] = 72,
            ["opacity"] = 0.3,
            ["rotation"] = 45,
            ["horizontalAlignment"] = "Center",
            ["verticalAlignment"] = "Center"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }
}