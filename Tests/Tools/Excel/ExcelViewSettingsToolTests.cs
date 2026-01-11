using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelViewSettingsTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelViewSettingsToolTests : ExcelTestBase
{
    private readonly ExcelViewSettingsTool _tool;

    public ExcelViewSettingsToolTests()
    {
        _tool = new ExcelViewSettingsTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void SetZoom_ShouldSetZoomLevel()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_zoom.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_zoom_output.xlsx");
        _tool.Execute("set_zoom", workbookPath, outputPath: outputPath, zoom: 150);
        var workbook = new Workbook(outputPath);
        Assert.Equal(150, workbook.Worksheets[0].Zoom);
    }

    [Fact]
    public void SetGridlines_ShouldSetGridlinesVisibility()
    {
        var workbookPath = CreateExcelWorkbook("test_set_gridlines.xlsx");
        var outputPath = CreateTestFilePath("test_set_gridlines_output.xlsx");
        _tool.Execute("set_gridlines", workbookPath, outputPath: outputPath, visible: false);
        var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].IsGridlinesVisible);
    }

    [Fact]
    public void SetTabColor_ShouldSetTabColor()
    {
        var workbookPath = CreateExcelWorkbook("test_set_tab_color.xlsx");
        var outputPath = CreateTestFilePath("test_set_tab_color_output.xlsx");
        _tool.Execute("set_tab_color", workbookPath, outputPath: outputPath, color: "FF0000");
        var workbook = new Workbook(outputPath);
        var tabColor = workbook.Worksheets[0].TabColor.ToArgb() & 0xFFFFFF;
        Assert.Equal(0xFF0000, tabColor);
    }

    [Fact]
    public void FreezePanes_ShouldFreezePanes()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_output.xlsx");
        var result = _tool.Execute("freeze_panes", workbookPath, outputPath: outputPath,
            freezeRow: 2, freezeColumn: 1);
        Assert.StartsWith("Panes frozen", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SET_ZOOM")]
    [InlineData("Set_Zoom")]
    [InlineData("set_zoom")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation.Replace("_", "")}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.xlsx");
        _tool.Execute(operation, workbookPath, outputPath: outputPath, zoom: 120);
        var workbook = new Workbook(outputPath);
        Assert.Equal(120, workbook.Worksheets[0].Zoom);
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
    public void SetZoom_WithSessionId_ShouldSetZoomInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_zoom.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("set_zoom", sessionId: sessionId, zoom: 150);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(150, workbook.Worksheets[0].Zoom);
    }

    [Fact]
    public void SetGridlines_WithSessionId_ShouldSetGridlinesInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_gridlines.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("set_gridlines", sessionId: sessionId, visible: false);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.False(workbook.Worksheets[0].IsGridlinesVisible);
    }

    [Fact]
    public void FreezePanes_WithSessionId_ShouldFreezeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_freeze.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("freeze_panes", sessionId: sessionId, freezeRow: 2, freezeColumn: 1);
        Assert.StartsWith("Panes frozen", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("set_zoom", sessionId: "invalid_session", zoom: 100));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateExcelWorkbookWithData("test_session_file.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath2);

        _tool.Execute("set_zoom", workbookPath1, sessionId, zoom: 175);

        var workbookAfter = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(175, workbookAfter.Worksheets[0].Zoom);
    }

    #endregion
}
