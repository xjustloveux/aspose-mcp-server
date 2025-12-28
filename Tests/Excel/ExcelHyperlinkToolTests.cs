using System.Text.Json;
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Hyperlink added to A1", result);
        Assert.Contains("https://example.com", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Hyperlinks);
        Assert.Equal("https://example.com", worksheet.Hyperlinks[0].Address);
        Assert.Equal("Click here", worksheet.Hyperlinks[0].TextToDisplay);
    }

    [Fact]
    public async Task AddHyperlink_WithoutDisplayText_ShouldAddHyperlink()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_no_display.xlsx");
        var outputPath = CreateTestFilePath("test_add_no_display_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "B2",
            ["url"] = "https://test.com"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Hyperlink added to B2", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Single(worksheet.Hyperlinks);
    }

    [Fact]
    public async Task AddHyperlink_CellAlreadyHasHyperlink_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_add_existing.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://existing.com");
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["cell"] = "A1",
            ["url"] = "https://new.com"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("already has a hyperlink", exception.Message);
    }

    [Fact]
    public async Task GetHyperlinks_ShouldReturnAllHyperlinks()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_hyperlinks.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Hyperlinks.Add("A1", 1, 1, "https://test1.com");
            worksheet.Hyperlinks.Add("B2", 1, 1, "https://test2.com");
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(2, root.GetProperty("count").GetInt32());
        var items = root.GetProperty("items");
        Assert.Equal(2, items.GetArrayLength());

        var first = items[0];
        Assert.Equal("A1", first.GetProperty("cell").GetString());
        Assert.Equal("https://test1.com", first.GetProperty("url").GetString());
    }

    [Fact]
    public async Task GetHyperlinks_EmptyWorksheet_ShouldReturnEmptyResult()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No hyperlinks found", root.GetProperty("message").GetString());
    }

    [Fact]
    public async Task EditHyperlink_ByCell_ShouldModifyHyperlink()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_hyperlink.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://old.com");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1",
            ["url"] = "https://new.com",
            ["displayText"] = "New Link"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("edited", result);
        Assert.Contains("url=https://new.com", result);

        using var resultWorkbook = new Workbook(outputPath);
        var hyperlink = resultWorkbook.Worksheets[0].Hyperlinks[0];
        Assert.Equal("https://new.com", hyperlink.Address);
        Assert.Equal("New Link", hyperlink.TextToDisplay);
    }

    [Fact]
    public async Task EditHyperlink_ByIndex_ShouldModifyHyperlink()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_by_index.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://first.com");
            workbook.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://second.com");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_by_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["hyperlinkIndex"] = 1,
            ["url"] = "https://modified.com"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("edited", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("https://first.com", resultWorkbook.Worksheets[0].Hyperlinks[0].Address);
        Assert.Equal("https://modified.com", resultWorkbook.Worksheets[0].Hyperlinks[1].Address);
    }

    [Fact]
    public async Task EditHyperlink_NoCellOrIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_missing.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["url"] = "https://new.com"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Either 'hyperlinkIndex' or 'cell' is required", exception.Message);
    }

    [Fact]
    public async Task EditHyperlink_CellNotFound_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_not_found.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["cell"] = "Z99",
            ["url"] = "https://new.com"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("No hyperlink found at cell", exception.Message);
    }

    [Fact]
    public async Task DeleteHyperlink_ByCell_ShouldDeleteHyperlink()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_hyperlink.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://delete.com");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("deleted", result);
        Assert.Contains("0 hyperlinks remaining", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].Hyperlinks);
    }

    [Fact]
    public async Task DeleteHyperlink_ByIndex_ShouldDeleteHyperlink()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_by_index.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://first.com");
            workbook.Worksheets[0].Hyperlinks.Add("B2", 1, 1, "https://second.com");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_by_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["hyperlinkIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("deleted", result);
        Assert.Contains("1 hyperlinks remaining", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Single(resultWorkbook.Worksheets[0].Hyperlinks);
        Assert.Equal("https://second.com", resultWorkbook.Worksheets[0].Hyperlinks[0].Address);
    }

    [Fact]
    public async Task DeleteHyperlink_InvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Hyperlinks.Add("A1", 1, 1, "https://test.com");
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["hyperlinkIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_op.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "invalid",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task AddHyperlink_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["cell"] = "A1",
            ["url"] = "https://test.com"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}