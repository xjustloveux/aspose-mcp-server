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

    private string CreateWorkbookWithFilter(string fileName, string range = "A1:C5")
    {
        var path = CreateExcelWorkbookWithData(fileName);
        using var workbook = new Workbook(path);
        workbook.Worksheets[0].AutoFilter.Range = range;
        workbook.Save(path);
        return path;
    }

    private string CreateWorkbookWithFilterData(string fileName)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var ws = workbook.Worksheets[0];
        ws.Cells["A1"].Value = "Status";
        ws.Cells["A2"].Value = "Active";
        ws.Cells["A3"].Value = "Inactive";
        ws.Cells["A4"].Value = "Active";
        ws.Cells["A5"].Value = "Pending";
        workbook.Save(path);
        return path;
    }

    #region General

    [Fact]
    public void Apply_ShouldApplyAutoFilter()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_apply.xlsx");
        var outputPath = CreateTestFilePath("test_apply_output.xlsx");
        var result = _tool.Execute("apply", workbookPath, range: "A1:C5", outputPath: outputPath);
        Assert.Contains("Auto filter applied to range A1:C5", result);
        using var workbook = new Workbook(outputPath);
        Assert.False(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void Apply_WithSheetIndex_ShouldApplyToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_apply_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells["A1"].Value = "Header";
            wb.Worksheets[1].Cells["A2"].Value = "Data";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_apply_sheet_output.xlsx");
        var result = _tool.Execute("apply", workbookPath, sheetIndex: 1, range: "A1:A2", outputPath: outputPath);
        Assert.Contains("Auto filter applied to range A1:A2", result);
        using var workbook = new Workbook(outputPath);
        Assert.False(string.IsNullOrEmpty(workbook.Worksheets[1].AutoFilter.Range));
    }

    [Fact]
    public void Remove_ShouldRemoveAutoFilter()
    {
        var workbookPath = CreateWorkbookWithFilter("test_remove.xlsx");
        var outputPath = CreateTestFilePath("test_remove_output.xlsx");
        var result = _tool.Execute("remove", workbookPath, outputPath: outputPath);
        Assert.Contains("Auto filter removed", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void Remove_NoExistingFilter_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_remove_no_filter.xlsx", 3);
        var outputPath = CreateTestFilePath("test_remove_no_filter_output.xlsx");
        var result = _tool.Execute("remove", workbookPath, outputPath: outputPath);
        Assert.Contains("Auto filter removed", result);
    }

    [Fact]
    public void Filter_ByValue_ShouldApplyCriteria()
    {
        var workbookPath = CreateWorkbookWithFilterData("test_filter_value.xlsx");
        var outputPath = CreateTestFilePath("test_filter_value_output.xlsx");
        var result = _tool.Execute("filter", workbookPath, range: "A1:A5", columnIndex: 0, criteria: "Active",
            outputPath: outputPath);
        Assert.Contains("Filter applied to column 0", result);

        using var resultWorkbook = new Workbook(outputPath);
        var ws = resultWorkbook.Worksheets[0];
        Assert.False(ws.Cells.Rows[0].IsHidden, "Header row should be visible");
        Assert.False(ws.Cells.Rows[1].IsHidden, "Row with 'Active' should be visible");
        Assert.True(ws.Cells.Rows[2].IsHidden, "Row with 'Inactive' should be hidden");
        Assert.False(ws.Cells.Rows[3].IsHidden, "Row with 'Active' should be visible");
        Assert.True(ws.Cells.Rows[4].IsHidden, "Row with 'Pending' should be hidden");
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
        var result = _tool.Execute("filter", workbookPath, range: "A1:B5", columnIndex: 1, criteria: "100",
            filterOperator: "GreaterThan", outputPath: outputPath);
        Assert.Contains("Filter applied to column 1", result);

        using var resultWorkbook = new Workbook(outputPath);
        var ws = resultWorkbook.Worksheets[0];
        Assert.False(ws.Cells.Rows[0].IsHidden, "Header row should be visible");
        Assert.True(ws.Cells.Rows[1].IsHidden, "Row with Amount=50 should be hidden");
        Assert.False(ws.Cells.Rows[2].IsHidden, "Row with Amount=150 should be visible");
        Assert.True(ws.Cells.Rows[3].IsHidden, "Row with Amount=75 should be hidden");
        Assert.False(ws.Cells.Rows[4].IsHidden, "Row with Amount=200 should be visible");
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
        var result = _tool.Execute("filter", workbookPath, range: "A1:A5", columnIndex: 0, criteria: "Smith",
            filterOperator: "Contains", outputPath: outputPath);
        Assert.Contains("Filter applied to column 0", result);

        using var resultWorkbook = new Workbook(outputPath);
        var ws = resultWorkbook.Worksheets[0];
        Assert.False(ws.Cells.Rows[0].IsHidden, "Header row should be visible");
        Assert.False(ws.Cells.Rows[1].IsHidden, "Row with 'John Smith' should be visible");
        Assert.True(ws.Cells.Rows[2].IsHidden, "Row with 'Jane Doe' should be hidden");
        Assert.True(ws.Cells.Rows[3].IsHidden, "Row with 'Bob Johnson' should be hidden");
        Assert.False(ws.Cells.Rows[4].IsHidden, "Row with 'Alice Smith' should be visible");
    }

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
        var result = _tool.Execute("filter", workbookPath, range: "A1:B5", columnIndex: 0, criteria: "test",
            filterOperator: operatorType, outputPath: outputPath);
        Assert.Contains("Filter applied to column 0", result);
    }

    [Fact]
    public void GetStatus_WithFilter_ShouldReturnEnabled()
    {
        var workbookPath = CreateWorkbookWithFilter("test_get_status_enabled.xlsx");
        var result = _tool.Execute("get_status", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
        Assert.Contains("A1:C5", json.RootElement.GetProperty("filterRange").GetString());
    }

    [Fact]
    public void GetStatus_WithoutFilter_ShouldReturnDisabled()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_status_disabled.xlsx", 3);
        var result = _tool.Execute("get_status", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
        Assert.Contains("not enabled", json.RootElement.GetProperty("status").GetString());
    }

    [Fact]
    public void GetStatus_ShouldIncludeWorksheetName()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_status_name.xlsx", 3);
        var result = _tool.Execute("get_status", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("worksheetName", out _));
    }

    [Theory]
    [InlineData("APPLY")]
    [InlineData("Apply")]
    [InlineData("apply")]
    public void Operation_ShouldBeCaseInsensitive_Apply(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:C5", outputPath: outputPath);
        Assert.Contains("Auto filter applied to range A1:C5", result);
    }

    [Theory]
    [InlineData("REMOVE")]
    [InlineData("Remove")]
    [InlineData("remove")]
    public void Operation_ShouldBeCaseInsensitive_Remove(string operation)
    {
        var workbookPath = CreateWorkbookWithFilter($"test_case_remove_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_remove_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, outputPath: outputPath);
        Assert.Contains("Auto filter removed", result);
    }

    [Theory]
    [InlineData("GET_STATUS")]
    [InlineData("Get_Status")]
    [InlineData("get_status")]
    public void Operation_ShouldBeCaseInsensitive_GetStatus(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("\"isFilterEnabled\":", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Apply_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("apply", workbookPath, sheetIndex: 99, range: "A1:C5"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Apply_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_apply_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("apply", workbookPath));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void Filter_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("filter", workbookPath, columnIndex: 0, criteria: "test"));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void Filter_WithMissingCriteria_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_missing_criteria.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("filter", workbookPath, range: "A1:A5", columnIndex: 0));
        Assert.Contains("criteria", ex.Message.ToLower());
    }

    [Fact]
    public void Filter_WithInvalidOperator_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_filter_invalid_op.xlsx", 3, 2);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("filter", workbookPath, range: "A1:A3", columnIndex: 0, criteria: "test",
                filterOperator: "InvalidOperator"));
        Assert.Contains("Unsupported filter operator", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get_status", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_status"));
    }

    #endregion

    #region Session

    [Fact]
    public void Apply_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_apply.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("apply", sessionId: sessionId, range: "A1:C5");
        Assert.Contains("Auto filter applied to range A1:C5", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.False(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void Remove_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateWorkbookWithFilter("test_session_remove.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("remove", sessionId: sessionId);
        Assert.Contains("Auto filter removed", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void Filter_WithSessionId_ShouldApplyInMemory()
    {
        var workbookPath = CreateWorkbookWithFilterData("test_session_filter.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("filter", sessionId: sessionId, range: "A1:A5", columnIndex: 0, criteria: "Active");
        Assert.Contains("Filter applied to column 0", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.False(string.IsNullOrEmpty(workbook.Worksheets[0].AutoFilter.Range));
    }

    [Fact]
    public void GetStatus_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithFilter("test_session_get_status.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_status", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isFilterEnabled").GetBoolean());
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_status", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateExcelWorkbookWithData("test_path_file.xlsx");
        var sessionWorkbook = CreateWorkbookWithFilter("test_session_file.xlsx");
        using (var wb = new Workbook(sessionWorkbook))
        {
            wb.Worksheets[0].Name = "SessionSheet";
            wb.Save(sessionWorkbook);
        }

        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get_status", pathWorkbook, sessionId);
        Assert.Contains("SessionSheet", result);
    }

    #endregion
}