using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelFilterTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

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
    public void Filter_ByValue_ShouldApplyCriteria()
    {
        var workbookPath = CreateWorkbookWithFilterData("test_filter_value.xlsx");
        var outputPath = CreateTestFilePath("test_filter_value_output.xlsx");
        var result = _tool.Execute("filter", workbookPath, range: "A1:A5", columnIndex: 0, criteria: "Active",
            outputPath: outputPath);
        Assert.Contains("Filter applied to column 0", result);
        using var resultWorkbook = new Workbook(outputPath);
        var ws = resultWorkbook.Worksheets[0];
        Assert.False(ws.Cells.Rows[0].IsHidden);
        Assert.False(ws.Cells.Rows[1].IsHidden);
        Assert.True(ws.Cells.Rows[2].IsHidden);
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
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("APPLY")]
    [InlineData("Apply")]
    [InlineData("apply")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:C5", outputPath: outputPath);
        Assert.Contains("Auto filter applied to range A1:C5", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_status"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Apply_WithSessionId_ShouldModifyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_apply.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("apply", sessionId: sessionId, range: "A1:C5");
        Assert.Contains("Auto filter applied to range A1:C5", result);
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
