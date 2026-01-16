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

    #region Additional Operations

    [Fact]
    public void SetHeaders_ShouldSetRowColumnHeadersVisibility()
    {
        var workbookPath = CreateExcelWorkbook("test_set_headers.xlsx");
        var outputPath = CreateTestFilePath("test_set_headers_output.xlsx");
        _tool.Execute("set_headers", workbookPath, outputPath: outputPath, visible: false);
        var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].IsRowColumnHeadersVisible);
    }

    [Fact]
    public void SetZeroValues_ShouldSetZeroValuesVisibility()
    {
        var workbookPath = CreateExcelWorkbook("test_set_zero.xlsx");
        var outputPath = CreateTestFilePath("test_set_zero_output.xlsx");
        _tool.Execute("set_zero_values", workbookPath, outputPath: outputPath, visible: false);
        var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].DisplayZeros);
    }

    [Fact]
    public void SetColumnWidth_ShouldSetWidth()
    {
        var workbookPath = CreateExcelWorkbook("test_column_width.xlsx");
        var outputPath = CreateTestFilePath("test_column_width_output.xlsx");
        _tool.Execute("set_column_width", workbookPath, outputPath: outputPath, columnIndex: 0, width: 25);
        var workbook = new Workbook(outputPath);
        Assert.Equal(25, workbook.Worksheets[0].Cells.GetColumnWidth(0));
    }

    [Fact]
    public void SetRowHeight_ShouldSetHeight()
    {
        var workbookPath = CreateExcelWorkbook("test_row_height.xlsx");
        var outputPath = CreateTestFilePath("test_row_height_output.xlsx");
        _tool.Execute("set_row_height", workbookPath, outputPath: outputPath, rowIndex: 0, height: 30);
        var workbook = new Workbook(outputPath);
        Assert.Equal(30, workbook.Worksheets[0].Cells.GetRowHeight(0));
    }

    [Fact]
    public void AutoFitColumn_ShouldAutoFitColumn()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_autofit_col.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_autofit_col_output.xlsx");
        var result = _tool.Execute("auto_fit_column", workbookPath, outputPath: outputPath, columnIndex: 0);
        Assert.Contains("auto-fit", result.ToLower());
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void AutoFitRow_ShouldAutoFitRow()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_autofit_row.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_autofit_row_output.xlsx");
        var result = _tool.Execute("auto_fit_row", workbookPath, outputPath: outputPath, rowIndex: 0);
        Assert.Contains("auto-fit", result.ToLower());
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SplitWindow_ShouldSplitWindow()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_split.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_split_output.xlsx");
        var result = _tool.Execute("split_window", workbookPath, outputPath: outputPath, splitRow: 5, splitColumn: 2);
        Assert.Contains("split", result.ToLower());
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ShowFormulas_ShouldToggleFormulasDisplay()
    {
        var workbookPath = CreateExcelWorkbook("test_formulas.xlsx");
        var outputPath = CreateTestFilePath("test_formulas_output.xlsx");
        var result = _tool.Execute("show_formulas", workbookPath, outputPath: outputPath, visible: true);
        Assert.Contains("formula", result.ToLower());
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetAll_ShouldSetMultipleSettings()
    {
        var workbookPath = CreateExcelWorkbook("test_set_all.xlsx");
        var outputPath = CreateTestFilePath("test_set_all_output.xlsx");
        var result = _tool.Execute("set_all", workbookPath, outputPath: outputPath,
            zoom: 125, showGridlines: false, showRowColumnHeaders: false, showZeroValues: false);
        Assert.Contains("view settings", result.ToLower());
        var workbook = new Workbook(outputPath);
        Assert.Equal(125, workbook.Worksheets[0].Zoom);
    }

    [Fact]
    public void UnfreezePanes_ShouldRemoveFreeze()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_unfreeze_output.xlsx");
        _tool.Execute("freeze_panes", workbookPath, outputPath: outputPath, freezeRow: 2, freezeColumn: 1);
        var result = _tool.Execute("freeze_panes", outputPath, outputPath: outputPath, unfreeze: true);
        Assert.Contains("unfrozen", result.ToLower());
    }

    #endregion
}
