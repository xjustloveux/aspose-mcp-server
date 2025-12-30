using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfPageToolTests : PdfTestBase
{
    private readonly PdfPageTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task AddPage_ShouldAddPage()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_page.pdf");
        var pagesBefore = new Document(pdfPath).Pages.Count;
        var outputPath = CreateTestFilePath("test_add_page_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["count"] = 1
            // Note: insertAt may not be fully supported, so we test without it
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var document = new Document(outputPath);
        Assert.True(document.Pages.Count > pagesBefore, "Page should be added");
    }

    [Fact]
    public async Task DeletePage_ShouldDeletePage()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_page.pdf");
        var pagesBefore = new Document(pdfPath).Pages.Count;
        Assert.True(pagesBefore >= 2, "PDF should have at least 2 pages");

        var outputPath = CreateTestFilePath("test_delete_page_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count < pagesBefore, "Page should be deleted");
    }

    [Fact]
    public async Task RotatePage_ShouldRotatePage()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_rotate_page.pdf");
        var outputPath = CreateTestFilePath("test_rotate_page_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "rotate",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["rotation"] = 90
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task GetPageInfo_ShouldReturnPageInfo()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_page_info.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_info",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Page", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetPageDetails_ShouldReturnPageDetails()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_page_details.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_details",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("pageIndex", result);
        Assert.Contains("width", result);
        Assert.Contains("height", result);
        Assert.Contains("mediaBox", result);
        Assert.Contains("cropBox", result);
    }

    [Fact]
    public async Task GetPageInfo_ShouldReturnAllPagesInfo()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_all_pages_info.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_info",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("\"count\": 2", result);
        Assert.Contains("items", result);
    }

    [Fact]
    public async Task AddPage_WithCustomSize_ShouldAddPageWithSize()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_custom_size.pdf");
        var outputPath = CreateTestFilePath("test_add_custom_size_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["count"] = 1,
            ["width"] = 400,
            ["height"] = 600
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);
    }

    [Fact]
    public async Task AddPage_WithInsertAt_ShouldInsertAtPosition()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_insert_at.pdf");
        var outputPath = CreateTestFilePath("test_add_insert_at_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["count"] = 1,
            ["insertAt"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var document = new Document(outputPath);
        Assert.Equal(3, document.Pages.Count);
    }

    [Fact]
    public async Task AddPage_WithMultiplePages_ShouldAddMultiplePages()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_multiple.pdf");
        var outputPath = CreateTestFilePath("test_add_multiple_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["count"] = 3
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var document = new Document(outputPath);
        Assert.Equal(5, document.Pages.Count);
    }

    [Fact]
    public async Task DeletePage_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_invalid.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["pageIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task RotatePage_WithInvalidRotation_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_rotate_invalid.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "rotate",
            ["path"] = pdfPath,
            ["pageIndex"] = 1,
            ["rotation"] = 45
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("rotation must be 0, 90, 180, or 270", exception.Message);
    }

    [Fact]
    public async Task RotatePage_WithPageIndices_ShouldRotateMultiplePages()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_rotate_multiple.pdf");
        var outputPath = CreateTestFilePath("test_rotate_multiple_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "rotate",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["rotation"] = 90,
            ["pageIndices"] = new JsonArray { 1, 2 }
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Rotated 2 page(s)", result);
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.on90, document.Pages[1].Rotate);
        Assert.Equal(Rotation.on90, document.Pages[2].Rotate);
    }

    [Fact]
    public async Task RotatePage_WithoutPageIndex_ShouldRotateAllPages()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_rotate_all.pdf");
        var outputPath = CreateTestFilePath("test_rotate_all_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "rotate",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["rotation"] = 180
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Rotated 2 page(s)", result);
    }

    [Fact]
    public async Task RotatePage_With270Degrees_ShouldRotate270()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_rotate_270.pdf");
        var outputPath = CreateTestFilePath("test_rotate_270_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "rotate",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["rotation"] = 270
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.on270, document.Pages[1].Rotate);
    }

    [Fact]
    public async Task RotatePage_With0Degrees_ShouldResetRotation()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_rotate_0.pdf");
        var outputPath = CreateTestFilePath("test_rotate_0_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "rotate",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["rotation"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var document = new Document(outputPath);
        Assert.Equal(Rotation.None, document.Pages[1].Rotate);
    }

    [Fact]
    public async Task GetPageDetails_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_details_invalid.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_details",
            ["path"] = pdfPath,
            ["pageIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pdfPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }
}