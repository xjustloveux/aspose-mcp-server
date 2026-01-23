using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelDataOperationsTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelDataOperationsToolTests : ExcelTestBase
{
    private readonly ExcelDataOperationsTool _tool;

    public ExcelDataOperationsToolTests()
    {
        _tool = new ExcelDataOperationsTool(SessionManager);
    }

    private string CreateWorkbookForSort(string fileName, bool withHeader = false)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var ws = workbook.Worksheets[0];
        if (withHeader)
        {
            ws.Cells["A1"].Value = "Name";
            ws.Cells["A2"].Value = "C";
            ws.Cells["A3"].Value = "A";
            ws.Cells["A4"].Value = "B";
        }
        else
        {
            ws.Cells["A1"].Value = "C";
            ws.Cells["A2"].Value = "A";
            ws.Cells["A3"].Value = "B";
        }

        workbook.Save(path);
        return path;
    }

    private string CreateWorkbookWithNumericData(string fileName)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var ws = workbook.Worksheets[0];
        ws.Cells["A1"].Value = 10;
        ws.Cells["A2"].Value = 20;
        ws.Cells["A3"].Value = 30;
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Sort_ShouldSortRange()
    {
        var workbookPath = CreateWorkbookForSort("test_sort.xlsx");
        var outputPath = CreateTestFilePath("test_sort_output.xlsx");
        _tool.Execute("sort", workbookPath, range: "A1:A3", sortColumn: 0, ascending: true, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("A", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("B", workbook.Worksheets[0].Cells["A2"].Value);
        Assert.Equal("C", workbook.Worksheets[0].Cells["A3"].Value);
    }

    [Fact]
    public void FindReplace_ShouldReplaceText()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_find_replace.xlsx", 3);
        var outputPath = CreateTestFilePath("test_find_replace_output.xlsx");
        _tool.Execute("find_replace", workbookPath, findText: "R1C1", replaceText: "Replaced", outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Replaced", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void BatchWrite_WithObjectFormat_ShouldWriteMultipleValues()
    {
        var workbookPath = CreateExcelWorkbook("test_batch_write.xlsx");
        var outputPath = CreateTestFilePath("test_batch_write_output.xlsx");
        var data = JsonNode.Parse("{\"A1\":\"Value1\",\"B1\":\"Value2\",\"A2\":\"Value3\"}");
        _tool.Execute("batch_write", workbookPath, data: data, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Value1", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("Value2", workbook.Worksheets[0].Cells["B1"].Value);
        Assert.Equal("Value3", workbook.Worksheets[0].Cells["A2"].Value);
    }

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_content.xlsx", 3);
        var result = _tool.Execute("get_content", workbookPath, range: "A1:B2");
        var data = GetResultData<GetContentResult>(result);
        Assert.NotEmpty(data.Rows);
        Assert.Contains(data.Rows, r => r.Values.Any(v => v?.ToString() == "R1C1"));
    }

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var workbookPath = CreateWorkbookWithNumericData("test_get_statistics.xlsx");
        var result = _tool.Execute("get_statistics", workbookPath, range: "A1:A3");
        var data = GetResultData<GetStatisticsResult>(result);
        Assert.NotEmpty(data.Worksheets);
        Assert.NotNull(data.Worksheets[0].RangeStatistics);
        Assert.Equal(60, data.Worksheets[0].RangeStatistics!.Sum);
    }

    [Fact]
    public void GetUsedRange_ShouldReturnUsedRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_used_range.xlsx", 3);
        var result = _tool.Execute("get_used_range", workbookPath);
        var data = GetResultData<GetUsedRangeResult>(result);
        Assert.NotNull(data.Range);
        Assert.NotEmpty(data.WorksheetName);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SORT")]
    [InlineData("Sort")]
    [InlineData("sort")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateWorkbookForSort($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:A3", sortColumn: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Sorted range", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown_operation", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_content"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Sort_WithSessionId_ShouldSortInMemory()
    {
        var workbookPath = CreateWorkbookForSort("test_session_sort.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("sort", sessionId: sessionId, range: "A1:A3", sortColumn: 0, ascending: true);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("A", workbook.Worksheets[0].Cells["A1"].Value?.ToString());
        Assert.Equal("B", workbook.Worksheets[0].Cells["A2"].Value?.ToString());
        Assert.Equal("C", workbook.Worksheets[0].Cells["A3"].Value?.ToString());
    }

    [Fact]
    public void FindReplace_WithSessionId_ShouldReplaceInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_find_replace.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("find_replace", sessionId: sessionId, findText: "R1C1", replaceText: "SessionReplaced");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("SessionReplaced", workbook.Worksheets[0].Cells["A1"].Value?.ToString());
    }

    [Fact]
    public void BatchWrite_WithSessionId_ShouldWriteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_batch_write.xlsx");
        var sessionId = OpenSession(workbookPath);
        var batchData = JsonNode.Parse("{\"A1\":\"SessionValue1\",\"B1\":\"SessionValue2\"}");
        var result = _tool.Execute("batch_write", sessionId: sessionId, data: batchData);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Batch write completed", data.Message);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("SessionValue1", workbook.Worksheets[0].Cells["A1"].Value?.ToString());
        Assert.Equal("SessionValue2", workbook.Worksheets[0].Cells["B1"].Value?.ToString());
    }

    [Fact]
    public void GetContent_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_content.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_content", sessionId: sessionId, range: "A1:B2");
        var data = GetResultData<GetContentResult>(result);
        Assert.NotEmpty(data.Rows);
        Assert.Contains(data.Rows, r => r.Values.Any(v => v?.ToString() == "R1C1"));
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithNumericData("test_session_get_statistics.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId, range: "A1:A3");
        var data = GetResultData<GetStatisticsResult>(result);
        Assert.NotEmpty(data.Worksheets);
        Assert.NotNull(data.Worksheets[0].RangeStatistics);
        Assert.Equal(60, data.Worksheets[0].RangeStatistics!.Sum);
    }

    [Fact]
    public void GetUsedRange_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_used_range.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_used_range", sessionId: sessionId);
        var data = GetResultData<GetUsedRangeResult>(result);
        Assert.NotNull(data.Range);
        Assert.NotEmpty(data.WorksheetName);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_content", sessionId: "invalid_session", range: "A1:B2"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateExcelWorkbook("test_path_file.xlsx");
        var sessionWorkbook = CreateExcelWorkbook("test_session_file.xlsx");
        using (var wb = new Workbook(sessionWorkbook))
        {
            wb.Worksheets[0].Name = "SessionSheet";
            wb.Worksheets[0].Cells["A1"].Value = "SessionData";
            wb.Save(sessionWorkbook);
        }

        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get_used_range", pathWorkbook, sessionId);
        var data = GetResultData<GetUsedRangeResult>(result);
        Assert.Equal("SessionSheet", data.WorksheetName);
    }

    #endregion
}
