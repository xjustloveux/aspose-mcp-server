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

    #region General Tests

    [Fact]
    public void MergeCells_ShouldMergeRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_cells.xlsx", 3);
        var outputPath = CreateTestFilePath("test_merge_cells_output.xlsx");
        var result = _tool.Execute(
            "merge",
            workbookPath,
            range: "A1:C1",
            outputPath: outputPath);
        Assert.Contains("Range A1:C1 merged", result);
        Assert.Contains("1 rows x 3 columns", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Cells.MergedCells.Count > 0);
    }

    [Fact]
    public void MergeCells_MultipleRows_ShouldMerge()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_multi.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_merge_multi_output.xlsx");
        var result = _tool.Execute(
            "merge",
            workbookPath,
            range: "A1:B3",
            outputPath: outputPath);
        Assert.Contains("Range A1:B3 merged", result);
        Assert.Contains("3 rows x 2 columns", result);

        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void MergeCells_SingleCell_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_single.xlsx", 3);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "merge",
            workbookPath,
            range: "A1"));
        Assert.Contains("Cannot merge a single cell", exception.Message);
    }

    [Fact]
    public void MergeCells_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_invalid_sheet.xlsx", 3);
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "merge",
            workbookPath,
            sheetIndex: 99,
            range: "A1:C1"));
    }

    [Fact]
    public void UnmergeCells_ShouldUnmergeRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unmerge_cells.xlsx", 3);
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3);
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_unmerge_cells_output.xlsx");
        var result = _tool.Execute(
            "unmerge",
            workbookPath,
            range: "A1:C1",
            outputPath: outputPath);
        Assert.Contains("Range A1:C1 unmerged", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void UnmergeCells_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unmerge_invalid_sheet.xlsx", 3);
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "unmerge",
            workbookPath,
            sheetIndex: 99,
            range: "A1:C1"));
    }

    [Fact]
    public void GetMergedCells_ShouldReturnMergedCells()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_merged_cells.xlsx", 3);
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells[0, 0].Value = "Header";
            workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3);
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(1, root.GetProperty("count").GetInt32());
        var items = root.GetProperty("items");
        Assert.Equal(1, items.GetArrayLength());

        var firstItem = items[0];
        Assert.Equal("A1:C1", firstItem.GetProperty("range").GetString());
        Assert.Equal("A1", firstItem.GetProperty("startCell").GetString());
        Assert.Equal("C1", firstItem.GetProperty("endCell").GetString());
        Assert.Equal(1, firstItem.GetProperty("rowCount").GetInt32());
        Assert.Equal(3, firstItem.GetProperty("columnCount").GetInt32());
        Assert.Equal("Header", firstItem.GetProperty("value").GetString());
    }

    [Fact]
    public void GetMergedCells_EmptyWorksheet_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_no_merged.xlsx");
        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No merged cells found", root.GetProperty("message").GetString());
    }

    [Fact]
    public void GetMergedCells_MultipleMergedRanges_ShouldReturnAll()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_multi_merged.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3); // A1:C1
            workbook.Worksheets[0].Cells.Merge(2, 0, 2, 2); // A3:B4
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(2, root.GetProperty("count").GetInt32());
        Assert.Equal(2, root.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void GetMergedCells_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "get",
            workbookPath,
            sheetIndex: 99));
    }

    [Fact]
    public void MergeCells_WithSheetIndex_ShouldMergeCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_merge_sheet_index.xlsx", 3);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells[0, 0].Value = "Test";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_merge_sheet_index_output.xlsx");
        var result = _tool.Execute(
            "merge",
            workbookPath,
            sheetIndex: 1,
            range: "A1:C1",
            outputPath: outputPath);
        Assert.Contains("merged", result);

        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Cells.MergedCells);
        Assert.Single(workbook.Worksheets[1].Cells.MergedCells);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_op.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid",
            workbookPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void MergeCells_MissingRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_range.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "merge",
            workbookPath));
        Assert.Contains("range", exception.Message.ToLower());
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetMergedCells_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_merged.xlsx", 3);
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells[0, 0].Value = "Header";
            workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3);
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "get",
            sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
    }

    [Fact]
    public void MergeCells_WithSessionId_ShouldMergeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_merge.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "merge",
            sessionId: sessionId,
            range: "A1:C1");
        Assert.Contains("Range A1:C1 merged", result);

        // Verify in-memory workbook has the merged cells
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.MergedCells.Count > 0);
    }

    [Fact]
    public void UnmergeCells_WithSessionId_ShouldUnmergeInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_unmerge.xlsx", 3);
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells.Merge(0, 0, 1, 3);
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "unmerge",
            sessionId: sessionId,
            range: "A1:C1");
        Assert.Contains("Range A1:C1 unmerged", result);

        // Verify in-memory workbook has no merged cells
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(sessionWorkbook.Worksheets[0].Cells.MergedCells);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}