using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Results.Excel.Render;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelRenderTool.
///     ExcelRenderTool is ReadOnly=true, so no session modification tests are needed.
///     Focuses on file-based rendering and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelRenderToolTests : ExcelTestBase
{
    private readonly ExcelRenderTool _tool;

    public ExcelRenderToolTests()
    {
        _tool = new ExcelRenderTool(SessionManager);
    }

    private string CreateWorkbookWithChart(string fileName)
    {
        var path = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        for (var i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].Value = $"Cat{i + 1}";
            sheet.Cells[i, 1].Value = (i + 1) * 10;
        }

        sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void RenderSheet_ShouldRenderToImage()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_render_sheet.xlsx");
        var outputPath = CreateTestFilePath("test_render_sheet_output.png");
        var result = _tool.Execute("render_sheet", workbookPath, outputPath: outputPath);
        var data = GetResultData<RenderExcelResult>(result);
        Assert.Contains("render", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(data.PageCount >= 1);
        Assert.Equal("png", data.Format);
        Assert.True(data.OutputPaths.Count >= 1);
        Assert.True(File.Exists(data.OutputPaths[0]));
    }

    [Fact]
    public void RenderSheet_WithJpegFormat_ShouldRenderToJpeg()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_render_jpeg.xlsx");
        var outputPath = CreateTestFilePath("test_render_jpeg_output.jpg");
        var result = _tool.Execute("render_sheet", workbookPath, outputPath: outputPath, format: "jpeg");
        var data = GetResultData<RenderExcelResult>(result);
        Assert.Contains("render", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("jpeg", data.Format);
        Assert.True(data.OutputPaths.Count >= 1);
        Assert.True(File.Exists(data.OutputPaths[0]));
    }

    [Fact]
    public void RenderChart_ShouldRenderChartToImage()
    {
        var workbookPath = CreateWorkbookWithChart("test_render_chart.xlsx");
        var outputPath = CreateTestFilePath("test_render_chart_output.png");
        var result = _tool.Execute("render_chart", workbookPath, outputPath: outputPath, chartIndex: 0);
        var data = GetResultData<RenderExcelResult>(result);
        Assert.Contains("render", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(data.PageCount >= 1);
        Assert.True(data.OutputPaths.Count >= 1);
        Assert.True(File.Exists(data.OutputPaths[0]));
    }

    [Fact]
    public void RenderSheet_WithCustomDpi_ShouldRender()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_render_dpi.xlsx");
        var outputPath = CreateTestFilePath("test_render_dpi_output.png");
        var result = _tool.Execute("render_sheet", workbookPath, outputPath: outputPath, dpi: 300);
        var data = GetResultData<RenderExcelResult>(result);
        Assert.Contains("render", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(data.OutputPaths[0]));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("RENDER_SHEET")]
    [InlineData("Render_Sheet")]
    [InlineData("render_sheet")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.png");
        var result = _tool.Execute(operation, workbookPath, outputPath: outputPath);
        var data = GetResultData<RenderExcelResult>(result);
        Assert.Contains("render", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx");
        var outputPath = CreateTestFilePath("test_unknown_op_output.png");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath, outputPath: outputPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        var outputPath = CreateTestFilePath("test_no_path_output.png");
        Assert.ThrowsAny<Exception>(() => _tool.Execute("render_sheet", outputPath: outputPath));
    }

    #endregion

    #region Session-based Read Tests

    [Fact]
    public void RenderSheet_WithSession_ShouldRenderFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_render_sheet.xlsx");
        var sessionId = OpenSession(workbookPath);
        var outputPath = CreateTestFilePath("test_session_render_sheet_output.png");
        var result = _tool.Execute("render_sheet", sessionId: sessionId, outputPath: outputPath);
        var data = GetResultData<RenderExcelResult>(result);
        Assert.Contains("render", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(data.OutputPaths.Count >= 1);
        Assert.True(File.Exists(data.OutputPaths[0]));
    }

    [Fact]
    public void RenderChart_WithSession_ShouldRenderFromMemory()
    {
        var workbookPath = CreateWorkbookWithChart("test_session_render_chart.xlsx");
        var sessionId = OpenSession(workbookPath);
        var outputPath = CreateTestFilePath("test_session_render_chart_output.png");
        var result = _tool.Execute("render_chart", sessionId: sessionId, outputPath: outputPath, chartIndex: 0);
        var data = GetResultData<RenderExcelResult>(result);
        Assert.Contains("render", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(data.OutputPaths.Count >= 1);
        Assert.True(File.Exists(data.OutputPaths[0]));
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_invalid_session_output.png");
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("render_sheet", sessionId: "invalid_session", outputPath: outputPath));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateExcelWorkbook("test_path_file.xlsx");
        var sessionWorkbook = CreateExcelWorkbookWithData("test_session_file.xlsx");
        var sessionId = OpenSession(sessionWorkbook);
        var outputPath = CreateTestFilePath("test_prefer_session_output.png");
        var result = _tool.Execute("render_sheet", pathWorkbook, sessionId, outputPath);
        var data = GetResultData<RenderExcelResult>(result);
        Assert.True(data.OutputPaths.Count >= 1);
        Assert.True(File.Exists(data.OutputPaths[0]));
    }

    #endregion
}
