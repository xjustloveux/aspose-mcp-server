using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

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

    #region General

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
    public void Sort_Descending_ShouldSortDescending()
    {
        var workbookPath = CreateWorkbookForSort("test_sort_desc.xlsx");
        var outputPath = CreateTestFilePath("test_sort_desc_output.xlsx");
        _tool.Execute("sort", workbookPath, range: "A1:A3", sortColumn: 0, ascending: false, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("C", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("B", workbook.Worksheets[0].Cells["A2"].Value);
        Assert.Equal("A", workbook.Worksheets[0].Cells["A3"].Value);
    }

    [Fact]
    public void Sort_WithHasHeader_ShouldSkipHeaderRow()
    {
        var workbookPath = CreateWorkbookForSort("test_sort_header.xlsx", true);
        var outputPath = CreateTestFilePath("test_sort_header_output.xlsx");
        _tool.Execute("sort", workbookPath, range: "A1:A4", sortColumn: 0, ascending: true, hasHeader: true,
            outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Name", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("A", workbook.Worksheets[0].Cells["A2"].Value);
        Assert.Equal("B", workbook.Worksheets[0].Cells["A3"].Value);
        Assert.Equal("C", workbook.Worksheets[0].Cells["A4"].Value);
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
    public void FindReplace_WithSubstring_ShouldNotLoopInfinitely()
    {
        var workbookPath = CreateExcelWorkbook("test_find_replace_substring.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = "Apple";
            wb.Worksheets[0].Cells["A2"].Value = "Apple Pie";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_find_replace_substring_output.xlsx");
        var result = _tool.Execute("find_replace", workbookPath, findText: "Apple", replaceText: "AppleTree",
            outputPath: outputPath);
        Assert.Contains("2 replacements", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("AppleTree", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("AppleTree Pie", workbook.Worksheets[0].Cells["A2"].Value);
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
    public void BatchWrite_WithArrayFormat_ShouldWriteValues()
    {
        var workbookPath = CreateExcelWorkbook("test_batch_write_array.xlsx");
        var outputPath = CreateTestFilePath("test_batch_write_array_output.xlsx");
        var data = JsonNode.Parse("[{\"cell\":\"A1\",\"value\":\"Value1\"},{\"cell\":\"B1\",\"value\":\"Value2\"}]");
        _tool.Execute("batch_write", workbookPath, data: data, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("Value1", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("Value2", workbook.Worksheets[0].Cells["B1"].Value);
    }

    [Fact]
    public void BatchWrite_WithNullData_ShouldSucceedWithZeroCells()
    {
        var workbookPath = CreateExcelWorkbook("test_batch_write_null.xlsx");
        var outputPath = CreateTestFilePath("test_batch_write_null_output.xlsx");
        var result = _tool.Execute("batch_write", workbookPath, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Batch write completed", result);
        Assert.Contains("0 cells written", result); // Verify specific content
    }

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_content.xlsx", 3);
        var result = _tool.Execute("get_content", workbookPath, range: "A1:B2");
        Assert.NotEmpty(result);
        Assert.Contains("R1C1", result);
    }

    [Fact]
    public void GetContent_WithoutRange_ShouldReturnAllContent()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_content_all.xlsx", 3);
        var result = _tool.Execute("get_content", workbookPath);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var workbookPath = CreateWorkbookWithNumericData("test_get_statistics.xlsx");
        var result = _tool.Execute("get_statistics", workbookPath, range: "A1:A3");
        Assert.Contains("Sum", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("60", result);
    }

    [Fact]
    public void GetUsedRange_ShouldReturnUsedRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_used_range.xlsx", 3);
        var result = _tool.Execute("get_used_range", workbookPath);
        Assert.Contains("range", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("SORT")]
    [InlineData("Sort")]
    [InlineData("sort")]
    public void Operation_ShouldBeCaseInsensitive_Sort(string operation)
    {
        var workbookPath = CreateWorkbookForSort($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:A3", sortColumn: 0, outputPath: outputPath);
        Assert.StartsWith("Sorted range", result);
    }

    [Theory]
    [InlineData("FIND_REPLACE")]
    [InlineData("Find_Replace")]
    [InlineData("find_replace")]
    public void Operation_ShouldBeCaseInsensitive_FindReplace(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_fr_{operation}.xlsx", 3);
        var outputPath = CreateTestFilePath($"test_case_fr_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, findText: "R1C1", replaceText: "X", outputPath: outputPath);
        Assert.StartsWith("Replaced", result);
    }

    [Theory]
    [InlineData("GET_CONTENT")]
    [InlineData("Get_Content")]
    [InlineData("get_content")]
    public void Operation_ShouldBeCaseInsensitive_GetContent(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_gc_{operation}.xlsx", 3);
        var result = _tool.Execute(operation, workbookPath, range: "A1:B2");
        Assert.NotEmpty(result);
    }

    [Theory]
    [InlineData("GET_STATISTICS")]
    [InlineData("Get_Statistics")]
    [InlineData("get_statistics")]
    public void Operation_ShouldBeCaseInsensitive_GetStatistics(string operation)
    {
        var workbookPath = CreateWorkbookWithNumericData($"test_case_gs_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:A3");
        Assert.Contains("sum", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("GET_USED_RANGE")]
    [InlineData("Get_Used_Range")]
    [InlineData("get_used_range")]
    public void Operation_ShouldBeCaseInsensitive_GetUsedRange(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_gur_{operation}.xlsx", 3);
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("range", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown_operation", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Sort_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_sort_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("sort", workbookPath, sortColumn: 0));
        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void FindReplace_WithMissingFindText_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_fr_missing_find.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("find_replace", workbookPath, replaceText: "New"));
        Assert.Contains("findText is required", ex.Message);
    }

    [Fact]
    public void FindReplace_WithMissingReplaceText_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_fr_missing_replace.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("find_replace", workbookPath, findText: "Old"));
        Assert.Contains("replaceText is required", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get_content", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_content"));
    }

    #endregion

    #region Session

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
        var data = JsonNode.Parse("{\"A1\":\"SessionValue1\",\"B1\":\"SessionValue2\"}");
        var result = _tool.Execute("batch_write", sessionId: sessionId, data: data);
        Assert.StartsWith("Batch write completed", result);
        Assert.Contains("session", result); // Verify session was used
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
        Assert.Contains("R1C1", result);
    }

    [Fact]
    public void GetStatistics_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithNumericData("test_session_get_statistics.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_statistics", sessionId: sessionId, range: "A1:A3");
        Assert.Contains("sum", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetUsedRange_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_used_range.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_used_range", sessionId: sessionId);
        Assert.Contains("range", result, StringComparison.OrdinalIgnoreCase);
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
        Assert.Contains("SessionSheet", result);
    }

    #endregion
}