using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

[SupportedOSPlatform("windows")]
public class ExcelImageToolTests : ExcelTestBase
{
    private readonly ExcelImageTool _tool = new();

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(100, 100);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Red);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    [Fact]
    public async Task AddImage_ShouldAddImageToWorksheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_image.xlsx");
        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image added to cell A1", result);
        Assert.Contains("size:", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Pictures);
    }

    [Fact]
    public async Task AddImage_WithDimensions_ShouldSetDimensions()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_image_dimensions.xlsx");
        var imagePath = CreateTestImage("test_image2.png");
        var outputPath = CreateTestFilePath("test_add_image_dimensions_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1",
            ["width"] = 200,
            ["height"] = 150
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image added to cell A1", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Pictures);
        var picture = worksheet.Pictures[0];
        Assert.True(Math.Abs(picture.Width - 200) < 10,
            $"Image width should be approximately 200, got {picture.Width}");
    }

    [Fact]
    public async Task AddImage_WithKeepAspectRatio_ShouldMaintainRatio()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_image_aspect.xlsx");
        var imagePath = CreateTestImage("test_image_aspect.png");
        var outputPath = CreateTestFilePath("test_add_image_aspect_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1",
            ["width"] = 200,
            ["keepAspectRatio"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image added to cell A1", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Pictures);
        Assert.True(worksheet.Pictures[0].IsLockAspectRatio);
    }

    [Fact]
    public async Task AddImage_WithoutKeepAspectRatio_ShouldAllowDistortion()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_image_no_aspect.xlsx");
        var imagePath = CreateTestImage("test_image_no_aspect.png");
        var outputPath = CreateTestFilePath("test_add_image_no_aspect_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1",
            ["width"] = 200,
            ["height"] = 50,
            ["keepAspectRatio"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image added to cell A1", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Pictures);
        Assert.False(worksheet.Pictures[0].IsLockAspectRatio);
    }

    [Fact]
    public async Task AddImage_UnsupportedFormat_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_unsupported.xlsx");
        var invalidImagePath = CreateTestFilePath("test_image.txt");
        await File.WriteAllTextAsync(invalidImagePath, "not an image");

        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["imagePath"] = invalidImagePath,
            ["cell"] = "A1"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unsupported image format", exception.Message);
    }

    [Fact]
    public async Task AddImage_FileNotFound_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_not_found.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["imagePath"] = @"C:\nonexistent\image.png",
            ["cell"] = "A1"
        };

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task AddImage_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet.xlsx");
        var imagePath = CreateTestImage("test_image_sheet.png");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1",
            ["sheetIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetImages_ShouldReturnAllImages()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_images.xlsx");
        var imagePath = CreateTestImage("test_image3.png");

        var addOutputPath = CreateTestFilePath("test_get_images_added.xlsx");
        var addArguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = addOutputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1"
        };
        await _tool.ExecuteAsync(addArguments);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = addOutputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(1, root.GetProperty("count").GetInt32());
        var items = root.GetProperty("items");
        Assert.Equal(1, items.GetArrayLength());

        var firstImage = items[0];
        Assert.True(firstImage.TryGetProperty("name", out _));
        Assert.True(firstImage.TryGetProperty("alternativeText", out _));
        Assert.True(firstImage.TryGetProperty("imageType", out _));
        Assert.True(firstImage.TryGetProperty("isLockAspectRatio", out _));

        var location = firstImage.GetProperty("location");
        Assert.True(location.TryGetProperty("upperLeftCell", out _));
        Assert.True(location.TryGetProperty("lowerRightCell", out _));
    }

    [Fact]
    public async Task GetImages_EmptyWorksheet_ShouldReturnEmptyResult()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No images found", root.GetProperty("message").GetString());
    }

    [Fact]
    public async Task DeleteImage_ShouldDeleteImage()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_image.xlsx");
        var imagePath = CreateTestImage("test_image4.png");

        var addOutputPath = CreateTestFilePath("test_delete_image_added.xlsx");
        var addArguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = addOutputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1"
        };
        await _tool.ExecuteAsync(addArguments);

        var outputPath = CreateTestFilePath("test_delete_image_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = addOutputPath,
            ["outputPath"] = outputPath,
            ["imageIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image #0 deleted", result);
        Assert.Contains("0 images remaining", result);

        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public async Task DeleteImage_WithRemainingImages_ShouldShowReorderWarning()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_reorder.xlsx");
        var imagePath1 = CreateTestImage("test_image_r1.png");
        var imagePath2 = CreateTestImage("test_image_r2.png");

        // Add first image
        var add1Path = CreateTestFilePath("test_delete_reorder_1.xlsx");
        await _tool.ExecuteAsync(new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = add1Path,
            ["imagePath"] = imagePath1,
            ["cell"] = "A1"
        });

        // Add second image
        var add2Path = CreateTestFilePath("test_delete_reorder_2.xlsx");
        await _tool.ExecuteAsync(new JsonObject
        {
            ["operation"] = "add",
            ["path"] = add1Path,
            ["outputPath"] = add2Path,
            ["imagePath"] = imagePath2,
            ["cell"] = "C1"
        });

        // Delete first image
        var outputPath = CreateTestFilePath("test_delete_reorder_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = add2Path,
            ["outputPath"] = outputPath,
            ["imageIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image #0 deleted", result);
        Assert.Contains("1 images remaining", result);
        Assert.Contains("re-ordered", result);

        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public async Task DeleteImage_InvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["imageIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task DeleteImage_NegativeIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_negative.xlsx");
        var imagePath = CreateTestImage("test_image_neg.png");

        var addPath = CreateTestFilePath("test_delete_negative_added.xlsx");
        await _tool.ExecuteAsync(new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = addPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1"
        });

        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = addPath,
            ["imageIndex"] = -1
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_op.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "invalid",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task AddImage_WithSheetIndex_ShouldAddToCorrectSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_sheet_index.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var imagePath = CreateTestImage("test_image_sheet2.png");
        var outputPath = CreateTestFilePath("test_add_sheet_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "B2",
            ["sheetIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image added to cell B2", result);

        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Pictures);
        Assert.Single(workbook.Worksheets[1].Pictures);
    }

    [Fact]
    public async Task ExtractImage_ShouldExtractImageToFile()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_extract_image.xlsx");
        var imagePath = CreateTestImage("test_image_extract.png");

        var addOutputPath = CreateTestFilePath("test_extract_image_added.xlsx");
        await _tool.ExecuteAsync(new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = addOutputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1"
        });

        var exportPath = CreateTestFilePath("extracted_image.png");
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = addOutputPath,
            ["imageIndex"] = 0,
            ["exportPath"] = exportPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Image #0", result);
        Assert.Contains("extracted to:", result);
        Assert.Contains(exportPath, result);
        Assert.True(File.Exists(exportPath));
        Assert.True(new FileInfo(exportPath).Length > 0);
    }

    [Fact]
    public async Task ExtractImage_ToJpeg_ShouldExtractAsJpeg()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_extract_jpeg.xlsx");
        var imagePath = CreateTestImage("test_image_jpeg.png");

        var addOutputPath = CreateTestFilePath("test_extract_jpeg_added.xlsx");
        await _tool.ExecuteAsync(new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = addOutputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1"
        });

        var exportPath = CreateTestFilePath("extracted_image.jpg");
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = addOutputPath,
            ["imageIndex"] = 0,
            ["exportPath"] = exportPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("extracted to:", result);
        Assert.True(File.Exists(exportPath));
    }

    [Fact]
    public async Task ExtractImage_InvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_extract_invalid.xlsx");
        var exportPath = CreateTestFilePath("extracted_invalid.png");
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = workbookPath,
            ["imageIndex"] = 99,
            ["exportPath"] = exportPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task ExtractImage_UnsupportedFormat_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_extract_unsupported.xlsx");
        var imagePath = CreateTestImage("test_image_unsupported.png");

        var addOutputPath = CreateTestFilePath("test_extract_unsupported_added.xlsx");
        await _tool.ExecuteAsync(new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = addOutputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1"
        });

        var exportPath = CreateTestFilePath("extracted_image.xyz");
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = addOutputPath,
            ["imageIndex"] = 0,
            ["exportPath"] = exportPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unsupported export format", exception.Message);
    }

    [Fact]
    public async Task ExtractImage_WithSheetIndex_ShouldExtractFromCorrectSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_extract_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var imagePath = CreateTestImage("test_image_extract_sheet.png");
        var addOutputPath = CreateTestFilePath("test_extract_sheet_added.xlsx");
        await _tool.ExecuteAsync(new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = addOutputPath,
            ["imagePath"] = imagePath,
            ["cell"] = "A1",
            ["sheetIndex"] = 1
        });

        var exportPath = CreateTestFilePath("extracted_sheet.png");
        var arguments = new JsonObject
        {
            ["operation"] = "extract",
            ["path"] = addOutputPath,
            ["sheetIndex"] = 1,
            ["imageIndex"] = 0,
            ["exportPath"] = exportPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("extracted to:", result);
        Assert.True(File.Exists(exportPath));
    }
}