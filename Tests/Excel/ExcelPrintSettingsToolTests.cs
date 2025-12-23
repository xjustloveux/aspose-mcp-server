using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelPrintSettingsToolTests : ExcelTestBase
{
    private readonly ExcelPrintSettingsTool _tool = new();

    [Fact]
    public async Task SetPrintArea_ShouldSetPrintArea()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_area.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_print_area_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_area",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:D10"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.NotNull(worksheet.PageSetup.PrintArea);
    }

    [Fact]
    public async Task SetPageSetup_ShouldSetPageSetup()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_page_setup.xlsx");
        var outputPath = CreateTestFilePath("test_set_page_setup_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["orientation"] = "Landscape",
            ["paperSize"] = "A4"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(PageOrientationType.Landscape, worksheet.PageSetup.Orientation);
    }

    [Fact]
    public async Task SetPrintTitles_ShouldSetPrintTitles()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_titles.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_print_titles_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_titles",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["rows"] = "1:1",
            ["columns"] = "A:A"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }

    [Fact]
    public async Task SetAll_ShouldSetAllPrintSettings()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_print_settings.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_all_print_settings_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_all",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:D10",
            ["orientation"] = "Portrait",
            ["paperSize"] = "A4"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.NotNull(worksheet.PageSetup);
        Assert.Equal(PageOrientationType.Portrait, worksheet.PageSetup.Orientation);
    }
}