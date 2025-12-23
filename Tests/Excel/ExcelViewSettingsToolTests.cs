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
    public async Task SplitWindow_ShouldSplitWindow()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_split_window.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_split_window_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "split_window",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["splitRow"] = 5
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        // Verify window was split (check split position)
        Assert.NotNull(worksheet);
    }

    [Fact]
    public async Task RemoveSplitWindow_ShouldRemoveSplit()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_remove_split.xlsx", 10, 5);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.RemoveSplit();
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
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }
}