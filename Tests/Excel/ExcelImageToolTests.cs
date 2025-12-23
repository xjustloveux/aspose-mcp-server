using System.Drawing;
using System.Drawing.Imaging;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelImageToolTests : ExcelTestBase
{
    private readonly ExcelImageTool _tool = new();

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(1, 1);
        bitmap.SetPixel(0, 0, Color.Red);
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
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Pictures.Count > 0, "Worksheet should contain at least one image");
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
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Pictures.Count > 0, "Worksheet should contain an image");
        var picture = worksheet.Pictures[0];
        Assert.True(Math.Abs(picture.Width - 200) < 10,
            $"Image width should be approximately 200, got {picture.Width}");
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
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Image", result, StringComparison.OrdinalIgnoreCase);
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

        var workbookBefore = new Workbook(addOutputPath);
        var imagesBefore = workbookBefore.Worksheets[0].Pictures.Count;
        Assert.True(imagesBefore > 0, "Image should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_image_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = addOutputPath,
            ["outputPath"] = outputPath,
            ["imageIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var imagesAfter = workbook.Worksheets[0].Pictures.Count;
        Assert.True(imagesAfter < imagesBefore,
            $"Image should be deleted. Before: {imagesBefore}, After: {imagesAfter}");
    }
}