using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

[SupportedOSPlatform("windows")]
public class ExcelImageToolTests : ExcelTestBase
{
    private readonly ExcelImageTool _tool;

    public ExcelImageToolTests()
    {
        _tool = new ExcelImageTool(SessionManager);
    }

    private string CreateTestImage(string fileName)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(100, 100);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Red);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    #region General Tests

    [Fact]
    public void AddImage_ShouldAddImageToWorksheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add_image.xlsx");
        var imagePath = CreateTestImage("test_image.png");
        var outputPath = CreateTestFilePath("test_add_image_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: outputPath);
        Assert.Contains("Image added to cell A1", result);
        Assert.Contains("size:", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Pictures);
    }

    [Fact]
    public void AddImage_WithDimensions_ShouldSetDimensions()
    {
        var workbookPath = CreateExcelWorkbook("test_add_image_dimensions.xlsx");
        var imagePath = CreateTestImage("test_image2.png");
        var outputPath = CreateTestFilePath("test_add_image_dimensions_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            width: 200,
            height: 150,
            outputPath: outputPath);
        Assert.Contains("Image added to cell A1", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Pictures);
        var picture = worksheet.Pictures[0];
        Assert.True(Math.Abs(picture.Width - 200) < 10,
            $"Image width should be approximately 200, got {picture.Width}");
    }

    [Fact]
    public void AddImage_WithKeepAspectRatio_ShouldMaintainRatio()
    {
        var workbookPath = CreateExcelWorkbook("test_add_image_aspect.xlsx");
        var imagePath = CreateTestImage("test_image_aspect.png");
        var outputPath = CreateTestFilePath("test_add_image_aspect_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            width: 200,
            keepAspectRatio: true,
            outputPath: outputPath);
        Assert.Contains("Image added to cell A1", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Pictures);
        Assert.True(worksheet.Pictures[0].IsLockAspectRatio);
    }

    [Fact]
    public void AddImage_WithoutKeepAspectRatio_ShouldAllowDistortion()
    {
        var workbookPath = CreateExcelWorkbook("test_add_image_no_aspect.xlsx");
        var imagePath = CreateTestImage("test_image_no_aspect.png");
        var outputPath = CreateTestFilePath("test_add_image_no_aspect_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            width: 200,
            height: 50,
            keepAspectRatio: false,
            outputPath: outputPath);
        Assert.Contains("Image added to cell A1", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Pictures);
        Assert.False(worksheet.Pictures[0].IsLockAspectRatio);
    }

    [Fact]
    public void AddImage_UnsupportedFormat_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_unsupported.xlsx");
        var invalidImagePath = CreateTestFilePath("test_image.txt");
        File.WriteAllText(invalidImagePath, "not an image");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            imagePath: invalidImagePath,
            cell: "A1"));
        Assert.Contains("Unsupported image format", exception.Message);
    }

    [Fact]
    public void AddImage_FileNotFound_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_not_found.xlsx");
        Assert.Throws<FileNotFoundException>(() => _tool.Execute(
            "add",
            workbookPath,
            imagePath: @"C:\nonexistent\image.png",
            cell: "A1"));
    }

    [Fact]
    public void AddImage_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet.xlsx");
        var imagePath = CreateTestImage("test_image_sheet.png");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            sheetIndex: 99,
            imagePath: imagePath,
            cell: "A1"));
    }

    [Fact]
    public void GetImages_ShouldReturnAllImages()
    {
        var workbookPath = CreateExcelWorkbook("test_get_images.xlsx");
        var imagePath = CreateTestImage("test_image3.png");

        var addOutputPath = CreateTestFilePath("test_get_images_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addOutputPath);
        var result = _tool.Execute(
            "get",
            addOutputPath);
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
    public void GetImages_EmptyWorksheet_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No images found", root.GetProperty("message").GetString());
    }

    [Fact]
    public void DeleteImage_ShouldDeleteImage()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_image.xlsx");
        var imagePath = CreateTestImage("test_image4.png");

        var addOutputPath = CreateTestFilePath("test_delete_image_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addOutputPath);

        var outputPath = CreateTestFilePath("test_delete_image_output.xlsx");
        var result = _tool.Execute(
            "delete",
            addOutputPath,
            imageIndex: 0,
            outputPath: outputPath);
        Assert.Contains("Image #0 deleted", result);
        Assert.Contains("0 images remaining", result);

        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void DeleteImage_WithRemainingImages_ShouldShowReorderWarning()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_reorder.xlsx");
        var imagePath1 = CreateTestImage("test_image_r1.png");
        var imagePath2 = CreateTestImage("test_image_r2.png");

        // Add first image
        var add1Path = CreateTestFilePath("test_delete_reorder_1.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath1,
            cell: "A1",
            outputPath: add1Path);

        // Add second image
        var add2Path = CreateTestFilePath("test_delete_reorder_2.xlsx");
        _tool.Execute(
            "add",
            add1Path,
            imagePath: imagePath2,
            cell: "C1",
            outputPath: add2Path);

        // Delete first image
        var outputPath = CreateTestFilePath("test_delete_reorder_output.xlsx");
        var result = _tool.Execute(
            "delete",
            add2Path,
            imageIndex: 0,
            outputPath: outputPath);
        Assert.Contains("Image #0 deleted", result);
        Assert.Contains("1 images remaining", result);
        Assert.Contains("re-ordered", result);

        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void DeleteImage_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            workbookPath,
            imageIndex: 99));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void DeleteImage_NegativeIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_negative.xlsx");
        var imagePath = CreateTestImage("test_image_neg.png");

        var addPath = CreateTestFilePath("test_delete_negative_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addPath);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            addPath,
            imageIndex: -1));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void AddImage_WithSheetIndex_ShouldAddToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add_sheet_index.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var imagePath = CreateTestImage("test_image_sheet2.png");
        var outputPath = CreateTestFilePath("test_add_sheet_index_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            sheetIndex: 1,
            imagePath: imagePath,
            cell: "B2",
            outputPath: outputPath);
        Assert.Contains("Image added to cell B2", result);

        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Pictures);
        Assert.Single(workbook.Worksheets[1].Pictures);
    }

    [Fact]
    public void ExtractImage_ShouldExtractImageToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_extract_image.xlsx");
        var imagePath = CreateTestImage("test_image_extract.png");

        var addOutputPath = CreateTestFilePath("test_extract_image_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addOutputPath);

        var exportPath = CreateTestFilePath("extracted_image.png");
        var result = _tool.Execute(
            "extract",
            addOutputPath,
            imageIndex: 0,
            exportPath: exportPath);
        Assert.Contains("Image #0", result);
        Assert.Contains("extracted to:", result);
        Assert.Contains(exportPath, result);
        Assert.True(File.Exists(exportPath));
        Assert.True(new FileInfo(exportPath).Length > 0);
    }

    [Fact]
    public void ExtractImage_ToJpeg_ShouldExtractAsJpeg()
    {
        var workbookPath = CreateExcelWorkbook("test_extract_jpeg.xlsx");
        var imagePath = CreateTestImage("test_image_jpeg.png");

        var addOutputPath = CreateTestFilePath("test_extract_jpeg_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addOutputPath);

        var exportPath = CreateTestFilePath("extracted_image.jpg");
        var result = _tool.Execute(
            "extract",
            addOutputPath,
            imageIndex: 0,
            exportPath: exportPath);
        Assert.Contains("extracted to:", result);
        Assert.True(File.Exists(exportPath));
    }

    [Fact]
    public void ExtractImage_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_extract_invalid.xlsx");
        var exportPath = CreateTestFilePath("extracted_invalid.png");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "extract",
            workbookPath,
            imageIndex: 99,
            exportPath: exportPath));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void ExtractImage_UnsupportedFormat_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_extract_unsupported.xlsx");
        var imagePath = CreateTestImage("test_image_unsupported.png");

        var addOutputPath = CreateTestFilePath("test_extract_unsupported_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addOutputPath);

        var exportPath = CreateTestFilePath("extracted_image.xyz");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "extract",
            addOutputPath,
            imageIndex: 0,
            exportPath: exportPath));
        Assert.Contains("Unsupported export format", exception.Message);
    }

    [Fact]
    public void ExtractImage_WithSheetIndex_ShouldExtractFromCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_extract_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var imagePath = CreateTestImage("test_image_extract_sheet.png");
        var addOutputPath = CreateTestFilePath("test_extract_sheet_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            sheetIndex: 1,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addOutputPath);

        var exportPath = CreateTestFilePath("extracted_sheet.png");
        var result = _tool.Execute(
            "extract",
            addOutputPath,
            sheetIndex: 1,
            imageIndex: 0,
            exportPath: exportPath);
        Assert.Contains("extracted to:", result);
        Assert.True(File.Exists(exportPath));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_op.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid",
            workbookPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void AddImage_MissingImagePath_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_image_path.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            cell: "A1"));
        Assert.Contains("imagepath", exception.Message.ToLower());
    }

    [Fact]
    public void AddImage_MissingCell_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_cell.xlsx");
        var imagePath = CreateTestImage("test_missing_cell_image.png");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath));
        Assert.Contains("cell", exception.Message.ToLower());
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetImages_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_images.xlsx");
        var imagePath = CreateTestImage("test_session_image.png");

        // Add image to file first
        var addOutputPath = CreateTestFilePath("test_session_get_images_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addOutputPath);

        var sessionId = OpenSession(addOutputPath);
        var result = _tool.Execute(
            "get",
            sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
    }

    [Fact]
    public void AddImage_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add_image.xlsx");
        var imagePath = CreateTestImage("test_session_add_image.png");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            imagePath: imagePath,
            cell: "B2");
        Assert.Contains("Image added to cell B2", result);

        // Verify in-memory workbook has the image
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void DeleteImage_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_delete_image.xlsx");
        var imagePath = CreateTestImage("test_session_delete_image.png");

        // Add image to file first
        var addOutputPath = CreateTestFilePath("test_session_delete_image_added.xlsx");
        _tool.Execute(
            "add",
            workbookPath,
            imagePath: imagePath,
            cell: "A1",
            outputPath: addOutputPath);

        var sessionId = OpenSession(addOutputPath);
        var result = _tool.Execute(
            "delete",
            sessionId: sessionId,
            imageIndex: 0);
        Assert.Contains("Image #0 deleted", result);

        // Verify in-memory workbook has no images
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}