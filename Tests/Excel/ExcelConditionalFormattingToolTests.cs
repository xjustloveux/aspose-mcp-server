using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelConditionalFormattingToolTests : ExcelTestBase
{
    private readonly ExcelConditionalFormattingTool _tool = new();

    [Fact]
    public async Task AddConditionalFormatting_ShouldAddFormatting()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_add_conditional_formatting.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_conditional_formatting_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A5",
            ["condition"] = "GreaterThan",
            ["value"] = "10"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.ConditionalFormattings.Count > 0, "Conditional formatting should be added");
    }

    [Fact]
    public async Task GetConditionalFormatting_ShouldReturnFormatting()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_conditional_formatting.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["formattingIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public async Task DeleteConditionalFormatting_ShouldDeleteFormatting()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_conditional_formatting.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);

        var formatCountBefore = worksheet.ConditionalFormattings.Count;
        Assert.True(formatCountBefore > 0, "Conditional formatting should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_conditional_formatting_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["conditionalFormattingIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var formatCountAfter = resultWorksheet.ConditionalFormattings.Count;
        Assert.True(formatCountAfter < formatCountBefore,
            $"Conditional formatting should be deleted. Before: {formatCountBefore}, After: {formatCountAfter}");
    }

    [Fact]
    public async Task EditConditionalFormatting_ShouldEditFormatting()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_edit_conditional_formatting.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:A5");
        var index = worksheet.ConditionalFormattings.Add();
        var formatting = worksheet.ConditionalFormattings[index];
        var area = new CellArea
        {
            StartRow = range.FirstRow, StartColumn = range.FirstColumn, EndRow = range.FirstRow + range.RowCount - 1,
            EndColumn = range.FirstColumn + range.ColumnCount - 1
        };
        formatting.AddArea(area);
        formatting.AddCondition(FormatConditionType.CellValue, OperatorType.GreaterThan, "10", null);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_conditional_formatting_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["conditionalFormattingIndex"] = 0,
            ["conditionIndex"] = 0,
            ["condition"] = "LessThan",
            ["value"] = "20"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.ConditionalFormattings.Count > 0,
            "Conditional formatting should exist after editing");
    }

    [Fact]
    public async Task Add_WithInvalidRange_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_range.xlsx");
        var outputPath = CreateTestFilePath("test_invalid_range_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "InvalidRange",
            ["condition"] = "GreaterThan",
            ["value"] = "10"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid range format", ex.Message);
    }

    [Fact]
    public async Task Add_WithBetweenCondition_ShouldUseBothFormulas()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_between_condition.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_between_condition_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A5",
            ["condition"] = "Between",
            ["value"] = "10",
            ["formula2"] = "50"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.ConditionalFormattings.Count > 0, "Conditional formatting should be added");
    }

    [Fact]
    public async Task Delete_WithInvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_invalid_index.xlsx");
        var outputPath = CreateTestFilePath("test_delete_invalid_index_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["conditionalFormattingIndex"] = 999
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public async Task Get_WithNoFormattings_ShouldReturnEmptyResult()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_no_formatting.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No conditional formattings found", result);
    }
}