using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Cells;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Image;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelImageTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddImageToWorksheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        var imagePath = CreateTestImage("test_add_image.png");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, imagePath: imagePath, cell: "A1", outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Delete_ShouldDeleteImage()
    {
        var workbookPath = CreateWorkbookWithImage("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, imageIndex: 0, outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Get_ShouldReturnAllImages()
    {
        var workbookPath = CreateWorkbookWithImage("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetImagesExcelResult>(result);
        Assert.Equal(1, data.Count);
    }

    [Fact]
    public void Extract_ShouldExtractImageToFile()
    {
        var workbookPath = CreateWorkbookWithImage("test_extract.xlsx");
        var exportPath = CreateTestFilePath("extracted.png");
        var result = _tool.Execute("extract", workbookPath, imageIndex: 0, exportPath: exportPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(exportPath));
        Assert.True(new FileInfo(exportPath).Length > 0);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var imagePath = CreateTestImage($"test_case_{operation}_image.png");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, imagePath: imagePath, cell: "A1", outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var imagePath = CreateTestImage("test_session_add_image.png");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, imagePath: imagePath, cell: "B2");
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithImage("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete", sessionId: sessionId, imageIndex: 0);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Pictures);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithImage("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetImagesExcelResult>(result);
        Assert.Equal(1, data.Count);
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
        var data = GetResultData<GetImagesExcelResult>(result);
        Assert.Equal(1, data.Count);
    }

    #endregion
}
