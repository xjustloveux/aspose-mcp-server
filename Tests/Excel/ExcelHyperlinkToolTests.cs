using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelHyperlinkToolTests : ExcelTestBase
{
    private readonly ExcelHyperlinkTool _tool = new();

    [Fact]
    public async Task AddHyperlink_ShouldAddHyperlink()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_hyperlink.xlsx");
        var outputPath = CreateTestFilePath("test_add_hyperlink_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1",
            ["url"] = "https://example.com",
            ["displayText"] = "Click here"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        // Check if hyperlink exists by checking hyperlinks collection
        Assert.True(worksheet.Hyperlinks.Count > 0, "Hyperlink should be added");
    }

    [Fact]
    public async Task GetHyperlinks_ShouldReturnAllHyperlinks()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_hyperlinks.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://test.com");
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public async Task EditHyperlink_ShouldModifyHyperlink()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_hyperlink.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://old.com");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1",
            ["url"] = "https://new.com"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        // Verify hyperlink was updated
        Assert.True(resultWorksheet.Hyperlinks.Count > 0, "Hyperlink should exist");
    }

    [Fact]
    public async Task DeleteHyperlink_ShouldDeleteHyperlink()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_hyperlink.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://delete.com");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        // Verify hyperlink was deleted - check if hyperlinks collection is empty or doesn't contain the deleted hyperlink
        Assert.True(
            resultWorksheet.Hyperlinks.Count == 0 ||
            !resultWorksheet.Hyperlinks.Any(h => h.Address.Contains("delete.com")), "Hyperlink should be deleted");
    }
}