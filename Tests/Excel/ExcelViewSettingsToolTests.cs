using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelViewSettingsToolTests : ExcelTestBase
{
    private readonly ExcelViewSettingsTool _tool = new();

    [Fact]
    public async Task SetZoom_ShouldSetZoomLevel()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_zoom.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_zoom_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_zoom",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["zoom"] = 150
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(150, worksheet.Zoom);
    }

    [Fact]
    public async Task SetGridlines_ShouldSetGridlinesVisibility()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_gridlines.xlsx");
        var outputPath = CreateTestFilePath("test_set_gridlines_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_gridlines",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["visible"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.False(worksheet.IsGridlinesVisible, "Gridlines should be hidden");
    }

    [Fact]
    public async Task SetColumnWidth_ShouldSetColumnWidth()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_column_width.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_column_width_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_column_width",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["columnIndex"] = 0,
            ["width"] = 20.0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(Math.Abs(worksheet.Cells.GetColumnWidth(0) - 20.0) < 0.1,
            $"Column width should be approximately 20, got {worksheet.Cells.GetColumnWidth(0)}");
    }

    [Fact]
    public async Task SetRowHeight_ShouldSetRowHeight()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_row_height.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_row_height_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_row_height",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["rowIndex"] = 0,
            ["height"] = 30.0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(Math.Abs(worksheet.Cells.GetRowHeight(0) - 30.0) < 0.1,
            $"Row height should be approximately 30, got {worksheet.Cells.GetRowHeight(0)}");
    }

    [Fact]
    public async Task SetHeaders_ShouldSetHeadersVisibility()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_headers.xlsx");
        var outputPath = CreateTestFilePath("test_set_headers_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_headers",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["visible"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.False(worksheet.IsRowColumnHeadersVisible, "Headers should be hidden");
    }

    [Fact]
    public async Task SetZeroValues_ShouldSetZeroValuesVisibility()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_zero_values.xlsx");
        var outputPath = CreateTestFilePath("test_set_zero_values_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_zero_values",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["visible"] = false
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.False(worksheet.DisplayZeros, "Zero values should be hidden");
    }

    [Fact]
    public async Task SetTabColor_ShouldSetTabColor()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_tab_color.xlsx");
        var outputPath = CreateTestFilePath("test_set_tab_color_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_tab_color",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["color"] = "FF0000" // Red
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var tabColor = worksheet.TabColor.ToArgb() & 0xFFFFFF;
        Assert.Equal(0xFF0000, tabColor);
    }

    [Fact]
    public async Task SetAll_ShouldSetMultipleSettings()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_view_settings.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_set_all_view_settings_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_all",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["zoom"] = 120,
            ["gridlinesVisible"] = true,
            ["headersVisible"] = true,
            ["zeroValuesVisible"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(120, worksheet.Zoom);
        Assert.True(worksheet.IsGridlinesVisible, "Gridlines should be visible");
        Assert.True(worksheet.IsRowColumnHeadersVisible, "Headers should be visible");
    }

    [Fact]
    public async Task FreezePanes_ShouldFreezePanes()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_freeze_panes.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_freeze_panes_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "freeze_panes",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["freezeRow"] = 2,
            ["freezeColumn"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("frozen", result);
    }

    [Fact]
    public async Task FreezePanes_Unfreeze_ShouldUnfreezePanes()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_unfreeze.xlsx", 10, 5);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.FreezePanes(2, 2, 2, 2);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_unfreeze_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "freeze_panes",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["unfreeze"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("unfrozen", result);
    }

    [Fact]
    public async Task AutoFitColumn_ShouldAutoFitColumn()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_auto_fit_column.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "This is a long text that needs auto fit";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_auto_fit_column_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "auto_fit_column",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["columnIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("auto-fitted", result);
    }

    [Fact]
    public async Task AutoFitColumn_WithRange_ShouldAutoFitColumnInRange()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_auto_fit_column_range.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_auto_fit_column_range_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "auto_fit_column",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["columnIndex"] = 0,
            ["startRow"] = 0,
            ["endRow"] = 5
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("auto-fitted", result);
    }

    [Fact]
    public async Task AutoFitRow_ShouldAutoFitRow()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_auto_fit_row.xlsx");
        var workbook = new Workbook(workbookPath);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        workbook.Worksheets[0].Cells["A1"].Value = "Line1\nLine2\nLine3";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_auto_fit_row_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "auto_fit_row",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["rowIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("auto-fitted", result);
    }

    [Fact]
    public async Task ShowFormulas_ShouldShowFormulas()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_show_formulas.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Formula = "=A1*2";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_show_formulas_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "show_formulas",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["visible"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("shown", result);
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(resultWorkbook.Worksheets[0].ShowFormulas);
    }

    [Fact]
    public async Task ShowFormulas_Hide_ShouldHideFormulas()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_hide_formulas.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].ShowFormulas = true;
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Formula = "=A1*2";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_hide_formulas_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "show_formulas",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["visible"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("hidden", result);
        var resultWorkbook = new Workbook(outputPath);
        Assert.False(resultWorkbook.Worksheets[0].ShowFormulas);
    }

    [Fact]
    public async Task FreezePanes_WithoutParams_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_freeze_no_params.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "freeze_panes",
            ["path"] = workbookPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task InvalidOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_operation.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "invalid_operation",
            ["path"] = workbookPath
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public async Task SetZoom_OutOfRange_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_zoom_out_of_range.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_zoom",
            ["path"] = workbookPath,
            ["zoom"] = 500 // Out of range (10-400)
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Zoom", ex.Message);
    }

    [Fact]
    public async Task SplitWindow_ShouldSplitWindow()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_split_window.xlsx", 20, 10);
        var outputPath = CreateTestFilePath("test_split_window_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "split_window",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["splitRow"] = 5,
            ["splitColumn"] = 3
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("split", result);
    }

    [Fact]
    public async Task SplitWindow_RowOnly_ShouldSplitHorizontally()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_split_row_only.xlsx", 20, 10);
        var outputPath = CreateTestFilePath("test_split_row_only_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "split_window",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["splitRow"] = 10
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("split", result);
    }

    [Fact]
    public async Task SplitWindow_RemoveSplit_ShouldRemoveSplit()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_remove_split.xlsx", 20, 10);
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].ActiveCell = "E10";
        workbook.Worksheets[0].Split();
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_remove_split_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "split_window",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["removeSplit"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        Assert.Contains("removed", result);
    }

    [Fact]
    public async Task SplitWindow_WithoutParams_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_split_no_params.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "split_window",
            ["path"] = workbookPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}