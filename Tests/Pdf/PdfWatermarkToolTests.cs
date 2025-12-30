using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfWatermarkToolTests : PdfTestBase
{
    private readonly PdfWatermarkTool _tool = new();

    private string CreatePdfDocument(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} Content"));
        }

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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 1 page(s)", result);
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added", result);
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
    public async Task AddWatermark_WithColor_ShouldApplyColor()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_color.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_color_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "URGENT",
            ["color"] = "Red"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added", result);
    }

    [Fact]
    public async Task AddWatermark_WithHexColor_ShouldApplyHexColor()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_hex.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_hex_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Custom Color",
            ["color"] = "#FF5500"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task AddWatermark_WithPageRange_ShouldApplyToSpecificPages()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_range.pdf", 5);
        var outputPath = CreateTestFilePath("test_add_watermark_range_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Selected Pages",
            ["pageRange"] = "1,3,5"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 3 page(s)", result);
    }

    [Fact]
    public async Task AddWatermark_WithPageRangeHyphen_ShouldApplyToRange()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_hyphen.pdf", 10);
        var outputPath = CreateTestFilePath("test_add_watermark_hyphen_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Range Pages",
            ["pageRange"] = "2-5"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 4 page(s)", result);
    }

    [Fact]
    public async Task AddWatermark_WithMixedPageRange_ShouldApplyCorrectly()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_mixed.pdf", 10);
        var outputPath = CreateTestFilePath("test_add_watermark_mixed_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Mixed Range",
            ["pageRange"] = "1,3-5,8"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 5 page(s)", result);
    }

    [Fact]
    public async Task AddWatermark_WithIsBackground_ShouldSetBackground()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_add_watermark_bg.pdf");
        var outputPath = CreateTestFilePath("test_add_watermark_bg_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Background Watermark",
            ["isBackground"] = true
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
        var pdfPath = CreatePdfDocument("test_add_watermark_all.pdf", 3);
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
            ["color"] = "Red",
            ["pageRange"] = "1-2",
            ["isBackground"] = true,
            ["horizontalAlignment"] = "Center",
            ["verticalAlignment"] = "Center"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 2 page(s)", result);
    }

    [Fact]
    public async Task AddWatermark_WithInvalidPageRange_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_invalid_range.pdf", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["text"] = "Test",
            ["pageRange"] = "invalid"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid page number", exception.Message);
    }

    [Fact]
    public async Task AddWatermark_WithOutOfBoundsPageRange_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_oob_range.pdf", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["text"] = "Test",
            ["pageRange"] = "1,5"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of bounds", exception.Message);
    }

    [Fact]
    public async Task AddWatermark_WithInvalidRangeFormat_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_invalid_format.pdf", 5);
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["text"] = "Test",
            ["pageRange"] = "3-1"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of bounds", exception.Message);
    }

    [Fact]
    public async Task Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_unknown_op.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pdfPath,
            ["text"] = "Test"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task AddWatermark_WithMultiplePages_ShouldApplyToAllPages()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_multi_page.pdf", 5);
        var outputPath = CreateTestFilePath("test_multi_page_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "All Pages"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
        Assert.Contains("Watermark added to 5 page(s)", result);
    }

    [Fact]
    public async Task AddWatermark_WithUnknownColor_ShouldUseGray()
    {
        // Arrange
        var pdfPath = CreatePdfDocument("test_unknown_color.pdf");
        var outputPath = CreateTestFilePath("test_unknown_color_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["text"] = "Unknown Color",
            ["color"] = "InvalidColor"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }
}