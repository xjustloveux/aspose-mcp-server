using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelMergeCellsToolTests : ExcelTestBase
{
    private readonly ExcelMergeCellsTool _tool;

    public ExcelMergeCellsToolTests()
    {
        _tool = new ExcelMergeCellsTool(SessionManager);
    }

    private string CreateWorkbookWithMergedCells(string fileName, string mergeRange = "A1:C1", string? value = "Header")
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        for (var row = 0; row < 5; row++)
        for (var col = 0; col < 5; col++)
            worksheet.Cells[row, col].Value = $"R{row}C{col}";
        if (!string.IsNullOrEmpty(value))
            worksheet.Cells[0, 0].Value = value;
        var parts = mergeRange.Replace(":", ",").Split(',');
        CellsHelper.CellNameToIndex(parts[0], out var startRow, out var startCol);
        var endPart = parts.Length > 1 ? parts[1] : parts[0];
        CellsHelper.CellNameToIndex(endPart, out var endRow, out var endCol);
        worksheet.Cells.Merge(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
        workbook.Save(workbookPath);
        return workbookPath;
    }

    #region General

    [Fact]
    public void Merge_ShouldMergeRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge.xlsx", 3);
        var outputPath = CreateTestFilePath("test_merge_output.xlsx");
        var result = _tool.Execute("merge", workbookPath, range: "A1:C1", outputPath: outputPath);
        Assert.StartsWith("Range A1:C1 merged", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.MergedCells.Count > 0);
    }

    [Fact]
    public void Merge_MultipleRows_ShouldMerge()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_multi.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_merge_multi_output.xlsx");
        var result = _tool.Execute("merge", workbookPath, range: "A1:B3", outputPath: outputPath);
        Assert.StartsWith("Range A1:B3 merged", result);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void Merge_WithSheetIndex_ShouldMergeCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_sheet.xlsx", 3);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells[0, 0].Value = "Test";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_merge_sheet_output.xlsx");
        var result = _tool.Execute("merge", workbookPath, sheetIndex: 1, range: "A1:C1", outputPath: outputPath);
        Assert.StartsWith("Range A1:C1 merged", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
        Assert.Single(workbook.Worksheets[1].Cells.MergedCells);
    }

    [Fact]
    public void Unmerge_ShouldUnmergeRange()
    {
        var workbookPath = CreateWorkbookWithMergedCells("test_unmerge.xlsx");
        var outputPath = CreateTestFilePath("test_unmerge_output.xlsx");
        var result = _tool.Execute("unmerge", workbookPath, range: "A1:C1", outputPath: outputPath);
        Assert.StartsWith("Range A1:C1 unmerged", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void Get_ShouldReturnMergedCellsInfo()
    {
        var workbookPath = CreateWorkbookWithMergedCells("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
        var firstItem = root.GetProperty("items")[0];
        Assert.Equal("A1:C1", firstItem.GetProperty("range").GetString());
        Assert.Equal("A1", firstItem.GetProperty("startCell").GetString());
        Assert.Equal("C1", firstItem.GetProperty("endCell").GetString());
        Assert.Equal(1, firstItem.GetProperty("rowCount").GetInt32());
        Assert.Equal(3, firstItem.GetProperty("columnCount").GetInt32());
        Assert.Equal("Header", firstItem.GetProperty("value").GetString());
    }

    [Fact]
    public void Get_Empty_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal("No merged cells found", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public void Get_MultipleMergedRanges_ShouldReturnAll()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_multi.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells.Merge(0, 0, 1, 3);
            wb.Worksheets[0].Cells.Merge(2, 0, 2, 2);
            wb.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(2, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Theory]
    [InlineData("MERGE")]
    [InlineData("Merge")]
    [InlineData("merge")]
    public void Operation_ShouldBeCaseInsensitive_Merge(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 3);
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:B1", outputPath: outputPath);
        Assert.StartsWith("Range A1:B1 merged", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("count", result);
    }

    [Theory]
    [InlineData("UNMERGE")]
    [InlineData("Unmerge")]
    [InlineData("unmerge")]
    public void Operation_ShouldBeCaseInsensitive_Unmerge(string operation)
    {
        var workbookPath = CreateWorkbookWithMergedCells($"test_case_unmerge_{operation}.xlsx", "A1:B1");
        var outputPath = CreateTestFilePath($"test_case_unmerge_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:B1", outputPath: outputPath);
        Assert.StartsWith("Range A1:B1 unmerged", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Merge_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_merge_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("merge", workbookPath));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void Merge_WithSingleCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_single.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("merge", workbookPath, range: "A1"));
        Assert.Contains("Cannot merge a single cell", ex.Message);
    }

    [Fact]
    public void Merge_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_invalid_sheet.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("merge", workbookPath, sheetIndex: 99, range: "A1:C1"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Unmerge_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unmerge_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unmerge", workbookPath));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void Unmerge_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unmerge_invalid_sheet.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unmerge", workbookPath, sheetIndex: 99, range: "A1:C1"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, sheetIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

    [Fact]
    public void Merge_WithSessionId_ShouldMergeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_merge.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("merge", sessionId: sessionId, range: "A1:C1");
        Assert.StartsWith("Range A1:C1 merged", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.MergedCells.Count > 0);
    }

    [Fact]
    public void Unmerge_WithSessionId_ShouldUnmergeInMemory()
    {
        var workbookPath = CreateWorkbookWithMergedCells("test_session_unmerge.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("unmerge", sessionId: sessionId, range: "A1:C1");
        Assert.StartsWith("Range A1:C1 unmerged", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithMergedCells("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateWorkbookWithMergedCells("test_session_file.xlsx", "A1:D1", "SessionMerged");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        Assert.Contains("SessionMerged", result);
    }

    #endregion
}