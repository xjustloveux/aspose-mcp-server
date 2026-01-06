using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelViewSettingsToolTests : ExcelTestBase
{
    private readonly ExcelViewSettingsTool _tool;

    public ExcelViewSettingsToolTests()
    {
        _tool = new ExcelViewSettingsTool(SessionManager);
    }

    #region General

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
    public void SetHeaders_ShouldSetHeadersVisibility()
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
        var workbookPath = CreateExcelWorkbook("test_set_zero_values.xlsx");
        var outputPath = CreateTestFilePath("test_set_zero_values_output.xlsx");
        _tool.Execute("set_zero_values", workbookPath, outputPath: outputPath, visible: false);
        var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].DisplayZeros);
    }

    [Fact]
    public void SetColumnWidth_ShouldSetColumnWidth()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_column_width.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_column_width_output.xlsx");
        _tool.Execute("set_column_width", workbookPath, outputPath: outputPath, columnIndex: 0, width: 20.0);
        var workbook = new Workbook(outputPath);
        Assert.True(Math.Abs(workbook.Worksheets[0].Cells.GetColumnWidth(0) - 20.0) < 0.1);
    }

    [Fact]
    public void SetRowHeight_ShouldSetRowHeight()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_row_height.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_row_height_output.xlsx");
        _tool.Execute("set_row_height", workbookPath, outputPath: outputPath, rowIndex: 0, height: 30.0);
        var workbook = new Workbook(outputPath);
        Assert.True(Math.Abs(workbook.Worksheets[0].Cells.GetRowHeight(0) - 30.0) < 0.1);
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
    public void SetAll_ShouldSetMultipleSettings()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_all.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_all_output.xlsx");
        _tool.Execute("set_all", workbookPath, outputPath: outputPath,
            zoom: 120, showGridlines: true, showRowColumnHeaders: true, showZeroValues: true);
        var workbook = new Workbook(outputPath);
        Assert.Equal(120, workbook.Worksheets[0].Zoom);
        Assert.True(workbook.Worksheets[0].IsGridlinesVisible);
        Assert.True(workbook.Worksheets[0].IsRowColumnHeadersVisible);
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

    [Fact]
    public void FreezePanes_Unfreeze_ShouldUnfreezePanes()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze.xlsx", 10, 5);
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].FreezePanes(2, 2, 2, 2);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_unfreeze_output.xlsx");
        var result = _tool.Execute("freeze_panes", workbookPath, outputPath: outputPath, unfreeze: true);
        Assert.StartsWith("Panes unfrozen", result);
    }

    [Fact]
    public void SplitWindow_ShouldSplitWindow()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_split.xlsx", 20, 10);
        var outputPath = CreateTestFilePath("test_split_output.xlsx");
        var result = _tool.Execute("split_window", workbookPath, outputPath: outputPath,
            splitRow: 5, splitColumn: 3);
        Assert.StartsWith("Window split", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SplitWindow_RemoveSplit_ShouldRemoveSplit()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_remove_split.xlsx", 20, 10);
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].ActiveCell = "E10";
        workbook.Worksheets[0].Split();
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_remove_split_output.xlsx");
        var result = _tool.Execute("split_window", workbookPath, outputPath: outputPath, removeSplit: true);
        Assert.Contains("split removed", result);
    }

    [Fact]
    public void AutoFitColumn_ShouldAutoFitColumn()
    {
        var workbookPath = CreateExcelWorkbook("test_auto_fit_column.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "This is a long text that needs auto fit";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_auto_fit_column_output.xlsx");
        var result = _tool.Execute("auto_fit_column", workbookPath, outputPath: outputPath, columnIndex: 0);
        Assert.StartsWith("Column 0 auto-fitted", result);
    }

    [Fact]
    public void AutoFitColumn_WithRange_ShouldAutoFitColumnInRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_auto_fit_column_range.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_auto_fit_column_range_output.xlsx");
        var result = _tool.Execute("auto_fit_column", workbookPath, outputPath: outputPath,
            columnIndex: 0, startRow: 0, endRow: 5);
        Assert.StartsWith("Column 0 auto-fitted", result);
    }

    [Fact]
    public void AutoFitRow_ShouldAutoFitRow()
    {
        var workbookPath = CreateExcelWorkbook("test_auto_fit_row.xlsx");
        var workbook = new Workbook(workbookPath);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        workbook.Worksheets[0].Cells["A1"].Value = "Line1\nLine2\nLine3";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_auto_fit_row_output.xlsx");
        var result = _tool.Execute("auto_fit_row", workbookPath, outputPath: outputPath, rowIndex: 0);
        Assert.StartsWith("Row 0 auto-fitted", result);
    }

    [Fact]
    public void ShowFormulas_Show_ShouldShowFormulas()
    {
        var workbookPath = CreateExcelWorkbook("test_show_formulas.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Formula = "=A1*2";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_show_formulas_output.xlsx");
        var result = _tool.Execute("show_formulas", workbookPath, outputPath: outputPath, visible: true);
        Assert.StartsWith("Formulas shown", result);
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(resultWorkbook.Worksheets[0].ShowFormulas);
    }

    [Fact]
    public void ShowFormulas_Hide_ShouldHideFormulas()
    {
        var workbookPath = CreateExcelWorkbook("test_hide_formulas.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].ShowFormulas = true;
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Formula = "=A1*2";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_hide_formulas_output.xlsx");
        var result = _tool.Execute("show_formulas", workbookPath, outputPath: outputPath, visible: false);
        Assert.StartsWith("Formulas hidden", result);
        var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.Worksheets[0].ShowFormulas);
    }

    [Theory]
    [InlineData("SET_ZOOM")]
    [InlineData("Set_Zoom")]
    [InlineData("set_zoom")]
    public void Operation_ShouldBeCaseInsensitive_SetZoom(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation.Replace("_", "")}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.xlsx");
        _tool.Execute(operation, workbookPath, outputPath: outputPath, zoom: 120);
        var workbook = new Workbook(outputPath);
        Assert.Equal(120, workbook.Worksheets[0].Zoom);
    }

    [Theory]
    [InlineData("SET_GRIDLINES")]
    [InlineData("Set_Gridlines")]
    [InlineData("set_gridlines")]
    public void Operation_ShouldBeCaseInsensitive_SetGridlines(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_grid_{operation.Replace("_", "")}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_grid_{operation.Replace("_", "")}_output.xlsx");
        _tool.Execute(operation, workbookPath, outputPath: outputPath, visible: false);
        var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].IsGridlinesVisible);
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
    public void SetZoom_OutOfRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_zoom_out_of_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_zoom", workbookPath, zoom: 500));
        Assert.Contains("Zoom", ex.Message);
    }

    [Fact]
    public void SetTabColor_WithMissingColor_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_tab_color_no_color.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_tab_color", workbookPath));
        Assert.Contains("color is required", ex.Message);
    }

    [Fact]
    public void SetBackground_WithMissingParams_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_background_no_params.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_background", workbookPath));
        Assert.Contains("imagePath or removeBackground", ex.Message);
    }

    [Fact]
    public void SetBackground_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var workbookPath = CreateExcelWorkbook("test_background_notfound.xlsx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("set_background", workbookPath, imagePath: "nonexistent.png"));
    }

    [Fact]
    public void FreezePanes_WithMissingParams_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_freeze_no_params.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("freeze_panes", workbookPath));
        Assert.Contains("freezeRow, freezeColumn, or unfreeze", ex.Message);
    }

    [Fact]
    public void SplitWindow_WithMissingParams_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_split_no_params.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("split_window", workbookPath));
        Assert.Contains("splitRow, splitColumn, or removeSplit", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("set_zoom", zoom: 100));
    }

    #endregion

    #region Session

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
    public void SetColumnWidth_WithSessionId_ShouldSetWidthInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_col_width.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("set_column_width", sessionId: sessionId, columnIndex: 0, width: 25.0);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(Math.Abs(workbook.Worksheets[0].Cells.GetColumnWidth(0) - 25.0) < 0.1);
    }

    [Fact]
    public void FreezePanes_WithSessionId_ShouldFreezeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_freeze.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("freeze_panes", sessionId: sessionId, freezeRow: 2, freezeColumn: 1);
        Assert.StartsWith("Panes frozen", result);
        Assert.Contains("session", result); // Verify session was used
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

        _ = SessionManager.GetDocument<Workbook>(sessionId);

        _tool.Execute("set_zoom", workbookPath1, sessionId, zoom: 175);

        var workbookAfter = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(175, workbookAfter.Worksheets[0].Zoom);
    }

    #endregion
}