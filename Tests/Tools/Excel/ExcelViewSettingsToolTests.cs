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

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_operation.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid_operation",
            workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    // Note: SetColumnWidth_WithMissingColumnIndex test removed - columnIndex has default value and is not nullable

    #endregion

    #region General Tests

    [Fact]
    public void SetZoom_ShouldSetZoomLevel()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_zoom.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_zoom_output.xlsx");
        _tool.Execute(
            "set_zoom",
            workbookPath,
            zoom: 150,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(150, worksheet.Zoom);
    }

    [Fact]
    public void SetGridlines_ShouldSetGridlinesVisibility()
    {
        var workbookPath = CreateExcelWorkbook("test_set_gridlines.xlsx");
        var outputPath = CreateTestFilePath("test_set_gridlines_output.xlsx");
        _tool.Execute(
            "set_gridlines",
            workbookPath,
            visible: false,
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.False(worksheet.IsGridlinesVisible, "Gridlines should be hidden");
    }

    [Fact]
    public void SetColumnWidth_ShouldSetColumnWidth()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_column_width.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_column_width_output.xlsx");
        _tool.Execute(
            "set_column_width",
            workbookPath,
            columnIndex: 0,
            width: 20.0,
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(Math.Abs(worksheet.Cells.GetColumnWidth(0) - 20.0) < 0.1,
            $"Column width should be approximately 20, got {worksheet.Cells.GetColumnWidth(0)}");
    }

    [Fact]
    public void SetRowHeight_ShouldSetRowHeight()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_row_height.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_row_height_output.xlsx");
        _tool.Execute(
            "set_row_height",
            workbookPath,
            rowIndex: 0,
            height: 30.0,
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(Math.Abs(worksheet.Cells.GetRowHeight(0) - 30.0) < 0.1,
            $"Row height should be approximately 30, got {worksheet.Cells.GetRowHeight(0)}");
    }

    [Fact]
    public void SetHeaders_ShouldSetHeadersVisibility()
    {
        var workbookPath = CreateExcelWorkbook("test_set_headers.xlsx");
        var outputPath = CreateTestFilePath("test_set_headers_output.xlsx");
        _tool.Execute(
            "set_headers",
            workbookPath,
            visible: false,
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.False(worksheet.IsRowColumnHeadersVisible, "Headers should be hidden");
    }

    [Fact]
    public void SetZeroValues_ShouldSetZeroValuesVisibility()
    {
        var workbookPath = CreateExcelWorkbook("test_set_zero_values.xlsx");
        var outputPath = CreateTestFilePath("test_set_zero_values_output.xlsx");
        _tool.Execute(
            "set_zero_values",
            workbookPath,
            visible: false,
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.False(worksheet.DisplayZeros, "Zero values should be hidden");
    }

    [Fact]
    public void SetTabColor_ShouldSetTabColor()
    {
        var workbookPath = CreateExcelWorkbook("test_set_tab_color.xlsx");
        var outputPath = CreateTestFilePath("test_set_tab_color_output.xlsx");
        _tool.Execute(
            "set_tab_color",
            workbookPath,
            color: "FF0000", // Red
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var tabColor = worksheet.TabColor.ToArgb() & 0xFFFFFF;
        Assert.Equal(0xFF0000, tabColor);
    }

    [Fact]
    public void SetAll_ShouldSetMultipleSettings()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_view_settings.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_all_view_settings_output.xlsx");
        _tool.Execute(
            "set_all",
            workbookPath,
            zoom: 120,
            showGridlines: true,
            showRowColumnHeaders: true,
            showZeroValues: true,
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(120, worksheet.Zoom);
        Assert.True(worksheet.IsGridlinesVisible, "Gridlines should be visible");
        Assert.True(worksheet.IsRowColumnHeadersVisible, "Headers should be visible");
    }

    [Fact]
    public void FreezePanes_ShouldFreezePanes()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_panes.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_panes_output.xlsx");
        var result = _tool.Execute(
            "freeze_panes",
            workbookPath,
            freezeRow: 2,
            freezeColumn: 1,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("frozen", result);
    }

    [Fact]
    public void FreezePanes_Unfreeze_ShouldUnfreezePanes()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze.xlsx", 10, 5);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.FreezePanes(2, 2, 2, 2);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_unfreeze_output.xlsx");
        var result = _tool.Execute(
            "freeze_panes",
            workbookPath,
            unfreeze: true,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("unfrozen", result);
    }

    [Fact]
    public void AutoFitColumn_ShouldAutoFitColumn()
    {
        var workbookPath = CreateExcelWorkbook("test_auto_fit_column.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "This is a long text that needs auto fit";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_auto_fit_column_output.xlsx");
        var result = _tool.Execute(
            "auto_fit_column",
            workbookPath,
            columnIndex: 0,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("auto-fitted", result);
    }

    [Fact]
    public void AutoFitColumn_WithRange_ShouldAutoFitColumnInRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_auto_fit_column_range.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_auto_fit_column_range_output.xlsx");
        var result = _tool.Execute(
            "auto_fit_column",
            workbookPath,
            columnIndex: 0,
            startRow: 0,
            endRow: 5,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("auto-fitted", result);
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
        var result = _tool.Execute(
            "auto_fit_row",
            workbookPath,
            rowIndex: 0,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("auto-fitted", result);
    }

    [Fact]
    public void ShowFormulas_ShouldShowFormulas()
    {
        var workbookPath = CreateExcelWorkbook("test_show_formulas.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Formula = "=A1*2";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_show_formulas_output.xlsx");
        var result = _tool.Execute(
            "show_formulas",
            workbookPath,
            visible: true,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("shown", result);
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
        var result = _tool.Execute(
            "show_formulas",
            workbookPath,
            visible: false,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("hidden", result);
        var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.Worksheets[0].ShowFormulas);
    }

    [Fact]
    public void FreezePanes_WithoutParams_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_freeze_no_params.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "freeze_panes",
            workbookPath));
    }

    [Fact]
    public void SetZoom_OutOfRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_zoom_out_of_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_zoom",
            workbookPath,
            zoom: 500)); // Out of range (10-400)
        Assert.Contains("Zoom", ex.Message);
    }

    [Fact]
    public void SplitWindow_ShouldSplitWindow()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_split_window.xlsx", 20, 10);
        var outputPath = CreateTestFilePath("test_split_window_output.xlsx");
        var result = _tool.Execute(
            "split_window",
            workbookPath,
            splitRow: 5,
            splitColumn: 3,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("split", result);
    }

    [Fact]
    public void SplitWindow_RowOnly_ShouldSplitHorizontally()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_split_row_only.xlsx", 20, 10);
        var outputPath = CreateTestFilePath("test_split_row_only_output.xlsx");
        var result = _tool.Execute(
            "split_window",
            workbookPath,
            splitRow: 10,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("split", result);
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
        var result = _tool.Execute(
            "split_window",
            workbookPath,
            removeSplit: true,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("removed", result);
    }

    [Fact]
    public void SplitWindow_WithoutParams_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_split_no_params.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "split_window",
            workbookPath));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void SetZoom_WithSessionId_ShouldSetZoomInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_set_zoom.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute(
            "set_zoom",
            sessionId: sessionId,
            zoom: 150);

        // Assert - verify in-memory change
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(150, workbook.Worksheets[0].Zoom);
    }

    [Fact]
    public void SetGridlines_WithSessionId_ShouldSetGridlinesInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_set_gridlines.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute(
            "set_gridlines",
            sessionId: sessionId,
            visible: false);

        // Assert - verify in-memory change
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.False(workbook.Worksheets[0].IsGridlinesVisible);
    }

    [Fact]
    public void SetColumnWidth_WithSessionId_ShouldSetWidthInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_set_col_width.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute(
            "set_column_width",
            sessionId: sessionId,
            columnIndex: 0,
            width: 25.0);

        // Assert - verify in-memory change
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(Math.Abs(workbook.Worksheets[0].Cells.GetColumnWidth(0) - 25.0) < 0.1);
    }

    [Fact]
    public void FreezePanes_WithSessionId_ShouldFreezeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_freeze.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "freeze_panes",
            sessionId: sessionId,
            freezeRow: 2,
            freezeColumn: 1);
        Assert.Contains("frozen", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute(
            "set_zoom",
            sessionId: "invalid_session_id",
            zoom: 100));
    }

    #endregion
}