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

    private string CreateWorkbookWithImage(string fileName, string cell = "A1")
    {
        var workbookPath = CreateExcelWorkbook(fileName);
        var imagePath = CreateTestImage($"img_for_{Path.GetFileNameWithoutExtension(fileName)}.png");
        var outputPath = CreateTestFilePath($"with_img_{fileName}");
        _tool.Execute("add", workbookPath, imagePath: imagePath, cell: cell, outputPath: outputPath);
        return outputPath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddImageToWorksheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var imagePath = CreateTestImage("test_add_image.png");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, imagePath: imagePath, cell: "A1", outputPath: outputPath);
        Assert.Contains("Image added to cell A1", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Add_WithDimensions_ShouldSetDimensions()
    {
        var workbookPath = CreateExcelWorkbook("test_add_dim.xlsx");
        var imagePath = CreateTestImage("test_add_dim_image.png");
        var outputPath = CreateTestFilePath("test_add_dim_output.xlsx");
        var result = _tool.Execute("add", workbookPath, imagePath: imagePath, cell: "A1",
            width: 200, height: 150, outputPath: outputPath);
        Assert.Contains("Image added to cell A1", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Pictures);
        var picture = workbook.Worksheets[0].Pictures[0];
        Assert.True(Math.Abs(picture.Width - 200) < 10);
    }

    [Fact]
    public void Add_WithKeepAspectRatio_ShouldMaintainRatio()
    {
        var workbookPath = CreateExcelWorkbook("test_add_aspect.xlsx");
        var imagePath = CreateTestImage("test_add_aspect_image.png");
        var outputPath = CreateTestFilePath("test_add_aspect_output.xlsx");
        var result = _tool.Execute("add", workbookPath, imagePath: imagePath, cell: "A1",
            width: 200, keepAspectRatio: true, outputPath: outputPath);
        Assert.Contains("Image added to cell A1", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Pictures[0].IsLockAspectRatio);
    }

    [Fact]
    public void Add_WithoutKeepAspectRatio_ShouldAllowDistortion()
    {
        var workbookPath = CreateExcelWorkbook("test_add_no_aspect.xlsx");
        var imagePath = CreateTestImage("test_add_no_aspect_image.png");
        var outputPath = CreateTestFilePath("test_add_no_aspect_output.xlsx");
        var result = _tool.Execute("add", workbookPath, imagePath: imagePath, cell: "A1",
            width: 200, height: 50, keepAspectRatio: false, outputPath: outputPath);
        Assert.Contains("Image added to cell A1", result);
        using var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].Pictures[0].IsLockAspectRatio);
    }

    [Fact]
    public void Add_WithSheetIndex_ShouldAddToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var imagePath = CreateTestImage("test_add_sheet_image.png");
        var outputPath = CreateTestFilePath("test_add_sheet_output.xlsx");
        var result = _tool.Execute("add", workbookPath, sheetIndex: 1, imagePath: imagePath,
            cell: "B2", outputPath: outputPath);
        Assert.Contains("Image added to cell B2", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Pictures);
        Assert.Single(workbook.Worksheets[1].Pictures);
    }

    [Fact]
    public void Delete_ShouldDeleteImage()
    {
        var workbookPath = CreateWorkbookWithImage("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, imageIndex: 0, outputPath: outputPath);
        Assert.Contains("Image #0 deleted", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Delete_WithRemainingImages_ShouldShowReorderWarning()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_reorder.xlsx");
        var imagePath1 = CreateTestImage("test_delete_r1.png");
        var imagePath2 = CreateTestImage("test_delete_r2.png");
        var add1Path = CreateTestFilePath("test_delete_r1.xlsx");
        _tool.Execute("add", workbookPath, imagePath: imagePath1, cell: "A1", outputPath: add1Path);
        var add2Path = CreateTestFilePath("test_delete_r2.xlsx");
        _tool.Execute("add", add1Path, imagePath: imagePath2, cell: "C1", outputPath: add2Path);
        var outputPath = CreateTestFilePath("test_delete_reorder_output.xlsx");
        var result = _tool.Execute("delete", add2Path, imageIndex: 0, outputPath: outputPath);
        Assert.Contains("Image #0 deleted", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Get_ShouldReturnAllImages()
    {
        var workbookPath = CreateWorkbookWithImage("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
        var firstImage = root.GetProperty("items")[0];
        Assert.True(firstImage.TryGetProperty("name", out _));
        Assert.True(firstImage.TryGetProperty("alternativeText", out _));
        Assert.True(firstImage.TryGetProperty("imageType", out _));
        Assert.True(firstImage.TryGetProperty("isLockAspectRatio", out _));
        var location = firstImage.GetProperty("location");
        Assert.True(location.TryGetProperty("upperLeftCell", out _));
        Assert.True(location.TryGetProperty("lowerRightCell", out _));
    }

    [Fact]
    public void Get_Empty_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal("No images found", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public void Extract_ShouldExtractImageToFile()
    {
        var workbookPath = CreateWorkbookWithImage("test_extract.xlsx");
        var exportPath = CreateTestFilePath("extracted.png");
        var result = _tool.Execute("extract", workbookPath, imageIndex: 0, exportPath: exportPath);
        Assert.Contains("Image #0", result);
        Assert.True(File.Exists(exportPath));
        Assert.True(new FileInfo(exportPath).Length > 0);
    }

    [Fact]
    public void Extract_ToJpeg_ShouldExtractAsJpeg()
    {
        var workbookPath = CreateWorkbookWithImage("test_extract_jpeg.xlsx");
        var exportPath = CreateTestFilePath("extracted.jpg");
        var result = _tool.Execute("extract", workbookPath, imageIndex: 0, exportPath: exportPath);
        Assert.Contains("Image #0", result);
        Assert.True(File.Exists(exportPath));
    }

    [Fact]
    public void Extract_WithSheetIndex_ShouldExtractFromCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_extract_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Save(workbookPath);
        }

        var imagePath = CreateTestImage("test_extract_sheet_image.png");
        var addPath = CreateTestFilePath("test_extract_sheet_add.xlsx");
        _tool.Execute("add", workbookPath, sheetIndex: 1, imagePath: imagePath, cell: "A1", outputPath: addPath);
        var exportPath = CreateTestFilePath("extracted_sheet.png");
        var result = _tool.Execute("extract", addPath, sheetIndex: 1, imageIndex: 0, exportPath: exportPath);
        Assert.Contains("Image #0", result);
        Assert.True(File.Exists(exportPath));
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var imagePath = CreateTestImage($"test_case_{operation}_image.png");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, imagePath: imagePath, cell: "A1", outputPath: outputPath);
        Assert.Contains("Image added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("\"count\":", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var workbookPath = CreateWorkbookWithImage($"test_case_del_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, imageIndex: 0, outputPath: outputPath);
        Assert.Contains("Image #0 deleted", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithMissingImagePath_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_path.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", workbookPath, cell: "A1"));
        Assert.Contains("imagepath", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_cell.xlsx");
        var imagePath = CreateTestImage("test_add_missing_cell_image.png");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", workbookPath, imagePath: imagePath));
        Assert.Contains("cell", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_unsupported.xlsx");
        var invalidPath = CreateTestFilePath("invalid.txt");
        File.WriteAllText(invalidPath, "not an image");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, imagePath: invalidPath, cell: "A1"));
        Assert.Contains("Unsupported image format", ex.Message);
    }

    [Fact]
    public void Add_WithFileNotFound_ShouldThrowFileNotFoundException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_notfound.xlsx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add", workbookPath, imagePath: @"C:\nonexistent\image.png", cell: "A1"));
    }

    [Fact]
    public void Add_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet.xlsx");
        var imagePath = CreateTestImage("test_add_invalid_sheet_image.png");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetIndex: 99, imagePath: imagePath, cell: "A1"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Delete_WithMissingImageIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_missing_index.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", workbookPath));
        Assert.Contains("imageIndex", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", workbookPath, imageIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Delete_WithNegativeIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithImage("test_delete_negative.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", workbookPath, imageIndex: -1));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", workbookPath, sheetIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Extract_WithMissingImageIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_extract_missing_index.xlsx");
        var exportPath = CreateTestFilePath("extract_missing.png");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", workbookPath, exportPath: exportPath));
        Assert.Contains("imageIndex", ex.Message);
    }

    [Fact]
    public void Extract_WithMissingExportPath_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithImage("test_extract_missing_path.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", workbookPath, imageIndex: 0));
        Assert.Contains("exportPath", ex.Message);
    }

    [Fact]
    public void Extract_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_extract_invalid.xlsx");
        var exportPath = CreateTestFilePath("extract_invalid.png");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", workbookPath, imageIndex: 99, exportPath: exportPath));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Extract_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithImage("test_extract_unsupported.xlsx");
        var exportPath = CreateTestFilePath("extract_unsupported.xyz");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("extract", workbookPath, imageIndex: 0, exportPath: exportPath));
        Assert.Contains("Unsupported export format", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var imagePath = CreateTestImage("test_session_add_image.png");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, imagePath: imagePath, cell: "B2");
        Assert.Contains("Image added to cell B2", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithImage("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete", sessionId: sessionId, imageIndex: 0);
        Assert.Contains("Image #0 deleted", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithImage("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Extract_WithSessionId_ShouldExtractFromMemory()
    {
        var workbookPath = CreateWorkbookWithImage("test_session_extract.xlsx");
        var sessionId = OpenSession(workbookPath);
        var exportPath = CreateTestFilePath("session_extracted.png");
        var result = _tool.Execute("extract", sessionId: sessionId, imageIndex: 0, exportPath: exportPath);
        Assert.Contains("Image #0", result);
        Assert.True(File.Exists(exportPath));
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateWorkbookWithImage("test_session_file.xlsx");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion
}