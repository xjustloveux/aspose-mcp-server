using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.MergeCells;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelMergeCellsTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

    [Fact]
    public void Merge_ShouldMergeRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge.xlsx", 3);
        var outputPath = CreateTestFilePath("test_merge_output.xlsx");
        var result = _tool.Execute("merge", workbookPath, range: "A1:C1", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Range A1:C1 merged", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.MergedCells.Count > 0);
    }

    [Fact]
    public void Unmerge_ShouldUnmergeRange()
    {
        var workbookPath = CreateWorkbookWithMergedCells("test_unmerge.xlsx");
        var outputPath = CreateTestFilePath("test_unmerge_output.xlsx");
        var result = _tool.Execute("unmerge", workbookPath, range: "A1:C1", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Range A1:C1 unmerged", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void Get_ShouldReturnMergedCellsInfo()
    {
        var workbookPath = CreateWorkbookWithMergedCells("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetMergedCellsResult>(result);
        Assert.Equal(1, data.Count);
    }

    [Fact]
    public void Get_Empty_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetMergedCellsResult>(result);
        Assert.Equal(0, data.Count);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("MERGE")]
    [InlineData("Merge")]
    [InlineData("merge")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 3);
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:B1", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Range A1:B1 merged", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Merge_WithSessionId_ShouldMergeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_merge.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("merge", sessionId: sessionId, range: "A1:C1");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Range A1:C1 merged", data.Message);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.MergedCells.Count > 0);
    }

    [Fact]
    public void Unmerge_WithSessionId_ShouldUnmergeInMemory()
    {
        var workbookPath = CreateWorkbookWithMergedCells("test_session_unmerge.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("unmerge", sessionId: sessionId, range: "A1:C1");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Range A1:C1 unmerged", data.Message);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithMergedCells("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetMergedCellsResult>(result);
        Assert.Equal(1, data.Count);
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
        var data = GetResultData<GetMergedCellsResult>(result);
        Assert.Contains(data.Items, item => item.Value == "SessionMerged");
    }

    #endregion
}
