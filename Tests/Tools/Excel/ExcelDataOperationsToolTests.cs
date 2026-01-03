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

    #region General Tests

    [Fact]
    public void SortData_ShouldSortRange()
    {
        var workbookPath = CreateExcelWorkbook("test_sort.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "C";
        worksheet.Cells["A2"].Value = "A";
        worksheet.Cells["A3"].Value = "B";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_sort_output.xlsx");
        _tool.Execute(
            "sort",
            workbookPath,
            range: "A1:A3",
            sortColumn: 0,
            ascending: true,
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("A", resultWorksheet.Cells["A1"].Value);
        Assert.Equal("B", resultWorksheet.Cells["A2"].Value);
        Assert.Equal("C", resultWorksheet.Cells["A3"].Value);
    }

    [Fact]
    public void FindReplace_ShouldReplaceText()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_find_replace.xlsx", 3);
        var outputPath = CreateTestFilePath("test_find_replace_output.xlsx");
        _tool.Execute(
            "find_replace",
            workbookPath,
            findText: "R1C1",
            replaceText: "Replaced",
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("Replaced", worksheet.Cells["A1"].Value);
    }

    [Fact]
    public void BatchWrite_ShouldWriteMultipleValues()
    {
        var workbookPath = CreateExcelWorkbook("test_batch_write.xlsx");
        var outputPath = CreateTestFilePath("test_batch_write_output.xlsx");
        var data = JsonNode.Parse("{\"A1\":\"Value1\",\"B1\":\"Value2\",\"A2\":\"Value3\"}");
        _tool.Execute(
            "batch_write",
            workbookPath,
            data: data,
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("Value1", worksheet.Cells["A1"].Value);
        Assert.Equal("Value2", worksheet.Cells["B1"].Value);
        Assert.Equal("Value3", worksheet.Cells["A2"].Value);
    }

    [Fact]
    public void GetContent_ShouldReturnContent()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_content.xlsx", 3);
        var result = _tool.Execute(
            "get_content",
            workbookPath,
            range: "A1:B2");
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("R1C1", result);
    }

    [Fact]
    public void GetStatistics_ShouldReturnStatistics()
    {
        var workbookPath = CreateExcelWorkbook("test_get_statistics.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = 10;
        worksheet.Cells["A2"].Value = 20;
        worksheet.Cells["A3"].Value = 30;
        workbook.Save(workbookPath);
        var result = _tool.Execute(
            "get_statistics",
            workbookPath,
            range: "A1:A3");
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Sum", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetUsedRange_ShouldReturnUsedRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_used_range.xlsx", 3);
        var result = _tool.Execute(
            "get_used_range",
            workbookPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Range", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SortData_WithHasHeader_ShouldSkipHeaderRow()
    {
        var workbookPath = CreateExcelWorkbook("test_sort_with_header.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["A2"].Value = "C";
        worksheet.Cells["A3"].Value = "A";
        worksheet.Cells["A4"].Value = "B";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_sort_with_header_output.xlsx");
        _tool.Execute(
            "sort",
            workbookPath,
            range: "A1:A4",
            sortColumn: 0,
            ascending: true,
            hasHeader: true,
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("Name", resultWorksheet.Cells["A1"].Value); // Header should remain
        Assert.Equal("A", resultWorksheet.Cells["A2"].Value);
        Assert.Equal("B", resultWorksheet.Cells["A3"].Value);
        Assert.Equal("C", resultWorksheet.Cells["A4"].Value);
    }

    [Fact]
    public void FindReplace_WithSubstring_ShouldNotLoopInfinitely()
    {
        // Arrange - Tests the fix for infinite loop when replaceText contains findText
        var workbookPath = CreateExcelWorkbook("test_find_replace_substring.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Apple";
        worksheet.Cells["A2"].Value = "Apple Pie";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_find_replace_substring_output.xlsx");

        // Act - Should complete without infinite loop
        var result = _tool.Execute(
            "find_replace",
            workbookPath,
            findText: "Apple",
            replaceText: "AppleTree",
            outputPath: outputPath);
        Assert.Contains("2 replacements", result);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("AppleTree", resultWorksheet.Cells["A1"].Value);
        Assert.Equal("AppleTree Pie", resultWorksheet.Cells["A2"].Value);
    }

    [Fact]
    public void BatchWrite_WithArrayFormat_ShouldWriteValues()
    {
        var workbookPath = CreateExcelWorkbook("test_batch_write_array.xlsx");
        var outputPath = CreateTestFilePath("test_batch_write_array_output.xlsx");
        var data = JsonNode.Parse("[{\"cell\":\"A1\",\"value\":\"Value1\"},{\"cell\":\"B1\",\"value\":\"Value2\"}]");
        _tool.Execute(
            "batch_write",
            workbookPath,
            data: data,
            outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("Value1", worksheet.Cells["A1"].Value);
        Assert.Equal("Value2", worksheet.Cells["B1"].Value);
    }

    [Fact]
    public void SortData_Descending_ShouldSortDescending()
    {
        var workbookPath = CreateExcelWorkbook("test_sort_desc.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "A";
        worksheet.Cells["A2"].Value = "C";
        worksheet.Cells["A3"].Value = "B";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_sort_desc_output.xlsx");
        _tool.Execute(
            "sort",
            workbookPath,
            range: "A1:A3",
            sortColumn: 0,
            ascending: false,
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("C", resultWorksheet.Cells["A1"].Value);
        Assert.Equal("B", resultWorksheet.Cells["A2"].Value);
        Assert.Equal("A", resultWorksheet.Cells["A3"].Value);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", workbookPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Sort_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_sort_missing_range.xlsx");
        var outputPath = CreateTestFilePath("test_sort_missing_range_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("sort", workbookPath, sortColumn: 0, outputPath: outputPath));

        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void FindReplace_WithMissingFindText_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_find_replace_missing.xlsx");
        var outputPath = CreateTestFilePath("test_find_replace_missing_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("find_replace", workbookPath, replaceText: "New", outputPath: outputPath));

        Assert.Contains("findText is required", ex.Message);
    }

    [Fact]
    public void BatchWrite_WithMissingData_ShouldSucceedWithZeroCells()
    {
        var workbookPath = CreateExcelWorkbook("test_batch_write_missing_data.xlsx");
        var outputPath = CreateTestFilePath("test_batch_write_missing_data_output.xlsx");

        // Act - When data is null, tool should succeed with 0 cells written
        var result = _tool.Execute("batch_write", workbookPath, outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Batch write completed", result);
        Assert.Contains("0 cells written", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetContent_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_content.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_content", sessionId: sessionId, range: "A1:B2");
        Assert.NotNull(result);
        Assert.Contains("R1C1", result);
    }

    [Fact]
    public void BatchWrite_WithSessionId_ShouldWriteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_batch_write.xlsx");
        var sessionId = OpenSession(workbookPath);
        var data = JsonNode.Parse("{\"A1\":\"SessionValue1\",\"B1\":\"SessionValue2\"}");
        var result = _tool.Execute("batch_write", sessionId: sessionId, data: data);
        Assert.Contains("Batch write completed", result);

        // Verify in-memory workbook has the values
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("SessionValue1", sessionWorkbook.Worksheets[0].Cells["A1"].Value?.ToString());
        Assert.Equal("SessionValue2", sessionWorkbook.Worksheets[0].Cells["B1"].Value?.ToString());
    }

    [Fact]
    public void Sort_WithSessionId_ShouldSortInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_sort.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "C";
        worksheet.Cells["A2"].Value = "A";
        worksheet.Cells["A3"].Value = "B";
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("sort", sessionId: sessionId, range: "A1:A3", sortColumn: 0, ascending: true);

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("A", sessionWorkbook.Worksheets[0].Cells["A1"].Value?.ToString());
        Assert.Equal("B", sessionWorkbook.Worksheets[0].Cells["A2"].Value?.ToString());
        Assert.Equal("C", sessionWorkbook.Worksheets[0].Cells["A3"].Value?.ToString());
    }

    [Fact]
    public void FindReplace_WithSessionId_ShouldReplaceInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_find_replace.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("find_replace", sessionId: sessionId, findText: "R1C1", replaceText: "SessionReplaced");

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("SessionReplaced", sessionWorkbook.Worksheets[0].Cells["A1"].Value?.ToString());
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_content", sessionId: "invalid_session_id", range: "A1:B2"));
    }

    #endregion
}