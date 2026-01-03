using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelFilterToolTests : ExcelTestBase
{
    private readonly ExcelFilterTool _tool;

    public ExcelFilterToolTests()
    {
        _tool = new ExcelFilterTool(SessionManager);
    }

    #region General Tests

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
    public void Filter_AllOperators_ShouldWork(string operatorType)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_op_{operatorType}.xlsx", 5, 2);
        var outputPath = CreateTestFilePath($"test_op_{operatorType}_output.xlsx");

        var result = _tool.Execute(
            "filter",
            workbookPath,
            range: "A1:B5",
            columnIndex: 0,
            criteria: "test",
            filterOperator: operatorType,
            outputPath: outputPath);

        Assert.Contains("Filter applied", result);
        Assert.Contains($"operator: {operatorType}", result);
    }

    #endregion

    #region Apply Tests

    [Fact]
    public void ApplyFilter_ShouldApplyAutoFilter()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_apply_filter.xlsx");
        var outputPath = CreateTestFilePath("test_apply_filter_output.xlsx");

        var result = _tool.Execute(
            "apply",
            workbookPath,
            range: "A1:C5",
            outputPath: outputPath);

        Assert.Contains("Auto filter applied", result);
        Assert.Contains("A1:C5", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.False(string.IsNullOrEmpty(worksheet.AutoFilter.Range));
    }

    [Fact]
    public void ApplyFilter_WithSheetIndex_ShouldApplyToCorrectSheet()
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

        var result = _tool.Execute(
            "apply",
            workbookPath,
            sheetIndex: 1,
            range: "A1:A2",
            outputPath: outputPath);

        Assert.Contains("sheet 1", result);
    }

    #endregion

    #region Remove Tests

    [Fact]
    public void RemoveFilter_ShouldRemoveAutoFilter()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_remove_filter.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].AutoFilter.Range = "A1:C5";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_remove_filter_output.xlsx");

        var result = _tool.Execute(
            "remove",
            workbookPath,
            outputPath: outputPath);

        Assert.Contains("Auto filter removed", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(resultWorkbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void RemoveFilter_NoExistingFilter_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_remove_no_filter.xlsx", 3);
        var outputPath = CreateTestFilePath("test_remove_no_filter_output.xlsx");

        var result = _tool.Execute(
            "remove",
            workbookPath,
            outputPath: outputPath);

        Assert.Contains("Auto filter removed", result);
    }

    #endregion

    #region Filter by Value Tests

    [Fact]
    public void Filter_ByValue_ShouldApplyCriteria()
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

        var result = _tool.Execute(
            "filter",
            workbookPath,
            range: "A1:A5",
            columnIndex: 0,
            criteria: "Active",
            outputPath: outputPath);

        Assert.Contains("Filter applied to column 0", result);
        Assert.Contains("criteria 'Active'", result);
    }

    [Fact]
    public void Filter_WithGreaterThanOperator_ShouldApplyCustomFilter()
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

        var result = _tool.Execute(
            "filter",
            workbookPath,
            range: "A1:B5",
            columnIndex: 1,
            criteria: "100",
            filterOperator: "GreaterThan",
            outputPath: outputPath);

        Assert.Contains("Filter applied to column 1", result);
        Assert.Contains("operator: GreaterThan", result);
    }

    [Fact]
    public void Filter_WithContainsOperator_ShouldApplyTextFilter()
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

        var result = _tool.Execute(
            "filter",
            workbookPath,
            range: "A1:A5",
            columnIndex: 0,
            criteria: "Smith",
            filterOperator: "Contains",
            outputPath: outputPath);

        Assert.Contains("Filter applied", result);
        Assert.Contains("operator: Contains", result);
    }

    [Fact]
    public void Filter_InvalidOperator_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_invalid_op.xlsx", 3, 2);

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "filter",
            workbookPath,
            range: "A1:A3",
            columnIndex: 0,
            criteria: "test",
            filterOperator: "InvalidOperator"));
        Assert.Contains("Unsupported filter operator", ex.Message);
    }

    #endregion

    #region Get Status Tests

    [Fact]
    public void GetFilterStatus_WithFilter_ShouldReturnEnabled()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_status_enabled.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].AutoFilter.Range = "A1:C5";
            wb.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get_status",
            workbookPath);

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
        Assert.Contains("A1:C5", json.RootElement.GetProperty("filterRange").GetString());
    }

    [Fact]
    public void GetFilterStatus_WithoutFilter_ShouldReturnDisabled()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_status_disabled.xlsx", 3);

        var result = _tool.Execute(
            "get_status",
            workbookPath);

        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
        Assert.Contains("not enabled", json.RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void GetFilterStatus_ShouldIncludeWorksheetName()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_status_name.xlsx", 3);

        var result = _tool.Execute(
            "get_status",
            workbookPath);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
    }

    #endregion

    #endregion

    #region Exception Tests

    [Fact]
    public void UnknownOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "unknown",
            workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_invalid_sheet.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "apply",
            workbookPath,
            sheetIndex: 99,
            range: "A1:C5"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Filter_MissingRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_missing_range.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "filter",
            workbookPath,
            columnIndex: 0,
            criteria: "test"));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void Apply_MissingRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_apply_missing_range.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "apply",
            workbookPath));
        Assert.Contains("range", ex.Message.ToLower());
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetStatus_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_status.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].AutoFilter.Range = "A1:C5";
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "get_status",
            sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
    }

    [Fact]
    public void Apply_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_apply.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "apply",
            sessionId: sessionId,
            range: "A1:C5");
        Assert.Contains("Auto filter applied", result);

        // Verify in-memory workbook has the filter
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.False(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void Remove_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_remove.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].AutoFilter.Range = "A1:C5";
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "remove",
            sessionId: sessionId);
        Assert.Contains("Auto filter removed", result);

        // Verify in-memory workbook has no filter
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void Filter_WithSessionId_ShouldApplyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_filter.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = "Status";
            wb.Worksheets[0].Cells["A2"].Value = "Active";
            wb.Worksheets[0].Cells["A3"].Value = "Inactive";
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "filter",
            sessionId: sessionId,
            range: "A1:A3",
            columnIndex: 0,
            criteria: "Active");
        Assert.Contains("Filter applied", result);

        // Verify in-memory workbook has the filter
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.False(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_status", sessionId: "invalid_session_id"));
    }

    #endregion
}