using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelFilterToolTests : ExcelTestBase
{
    private readonly ExcelFilterTool _tool = new();

    #region All Operators Test

    [Theory]
    [InlineData("Equal")]
    [InlineData("NotEqual")]
    [InlineData("GreaterThan")]
    [InlineData("GreaterOrEqual")]
    [InlineData("LessThan")]
    [InlineData("LessOrEqual")]
    [InlineData("Contains")]
    [InlineData("NotContains")]
    [InlineData("BeginsWith")]
    [InlineData("EndsWith")]
    public async Task Filter_AllOperators_ShouldWork(string operatorType)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_op_{operatorType}.xlsx", 5, 2);
        var outputPath = CreateTestFilePath($"test_op_{operatorType}_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "filter",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:B5",
            ["columnIndex"] = 0,
            ["criteria"] = "test",
            ["filterOperator"] = operatorType
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Filter applied", result);
        Assert.Contains($"operator: {operatorType}", result);
    }

    #endregion

    #region Apply Tests

    [Fact]
    public async Task ApplyFilter_ShouldApplyAutoFilter()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_apply_filter.xlsx");
        var outputPath = CreateTestFilePath("test_apply_filter_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:C5"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Auto filter applied", result);
        Assert.Contains("A1:C5", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.False(string.IsNullOrEmpty(worksheet.AutoFilter.Range));
    }

    [Fact]
    public async Task ApplyFilter_WithSheetIndex_ShouldApplyToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_apply_filter_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells["A1"].Value = "Header";
            wb.Worksheets[1].Cells["A2"].Value = "Data";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_apply_filter_sheet_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 1,
            ["range"] = "A1:A2"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("sheet 1", result);
    }

    #endregion

    #region Remove Tests

    [Fact]
    public async Task RemoveFilter_ShouldRemoveAutoFilter()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_remove_filter.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].AutoFilter.Range = "A1:C5";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_remove_filter_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "remove",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Auto filter removed", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(resultWorkbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public async Task RemoveFilter_NoExistingFilter_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_remove_no_filter.xlsx", 3);
        var outputPath = CreateTestFilePath("test_remove_no_filter_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "remove",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Auto filter removed", result);
    }

    #endregion

    #region Filter by Value Tests

    [Fact]
    public async Task Filter_ByValue_ShouldApplyCriteria()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_value.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = "Status";
            wb.Worksheets[0].Cells["A2"].Value = "Active";
            wb.Worksheets[0].Cells["A3"].Value = "Inactive";
            wb.Worksheets[0].Cells["A4"].Value = "Active";
            wb.Worksheets[0].Cells["A5"].Value = "Pending";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_filter_value_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "filter",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A5",
            ["columnIndex"] = 0,
            ["criteria"] = "Active"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Filter applied to column 0", result);
        Assert.Contains("criteria 'Active'", result);
    }

    [Fact]
    public async Task Filter_WithGreaterThanOperator_ShouldApplyCustomFilter()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_gt.xlsx", 5, 2);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["B1"].Value = "Amount";
            wb.Worksheets[0].Cells["B2"].Value = 50;
            wb.Worksheets[0].Cells["B3"].Value = 150;
            wb.Worksheets[0].Cells["B4"].Value = 75;
            wb.Worksheets[0].Cells["B5"].Value = 200;
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_filter_gt_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "filter",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:B5",
            ["columnIndex"] = 1,
            ["criteria"] = "100",
            ["filterOperator"] = "GreaterThan"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Filter applied to column 1", result);
        Assert.Contains("operator: GreaterThan", result);
    }

    [Fact]
    public async Task Filter_WithContainsOperator_ShouldApplyTextFilter()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_contains.xlsx", 5, 2);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = "Name";
            wb.Worksheets[0].Cells["A2"].Value = "John Smith";
            wb.Worksheets[0].Cells["A3"].Value = "Jane Doe";
            wb.Worksheets[0].Cells["A4"].Value = "Bob Johnson";
            wb.Worksheets[0].Cells["A5"].Value = "Alice Smith";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_filter_contains_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "filter",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A5",
            ["columnIndex"] = 0,
            ["criteria"] = "Smith",
            ["filterOperator"] = "Contains"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Filter applied", result);
        Assert.Contains("operator: Contains", result);
    }

    [Fact]
    public async Task Filter_InvalidOperator_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_invalid_op.xlsx", 3, 2);
        var arguments = new JsonObject
        {
            ["operation"] = "filter",
            ["path"] = workbookPath,
            ["range"] = "A1:A3",
            ["columnIndex"] = 0,
            ["criteria"] = "test",
            ["filterOperator"] = "InvalidOperator"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unsupported filter operator", ex.Message);
    }

    #endregion

    #region Get Status Tests

    [Fact]
    public async Task GetFilterStatus_WithFilter_ShouldReturnEnabled()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_status_enabled.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].AutoFilter.Range = "A1:C5";
            wb.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_status",
            ["path"] = workbookPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
        Assert.Contains("A1:C5", json.RootElement.GetProperty("filterRange").GetString());
    }

    [Fact]
    public async Task GetFilterStatus_WithoutFilter_ShouldReturnDisabled()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_status_disabled.xlsx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "get_status",
            ["path"] = workbookPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
        Assert.Contains("not enabled", json.RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public async Task GetFilterStatus_ShouldIncludeWorksheetName()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_status_name.xlsx", 3);
        var arguments = new JsonObject
        {
            ["operation"] = "get_status",
            ["path"] = workbookPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
    }

    #endregion

    #region Error Handling Tests

    [Fact]
    public async Task UnknownOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = workbookPath
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public async Task InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "apply",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["range"] = "A1:C5"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}