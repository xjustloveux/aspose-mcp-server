using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelGroupToolTests : ExcelTestBase
{
    private readonly ExcelGroupTool _tool;

    public ExcelGroupToolTests()
    {
        _tool = new ExcelGroupTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void GroupRows_ShouldGroupRows()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_rows_output.xlsx");
        var result = _tool.Execute(
            "group_rows",
            workbookPath,
            startRow: 1,
            endRow: 3,
            outputPath: outputPath);
        Assert.Contains("Rows 1-3 grouped", result);
        Assert.Contains(outputPath, result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Cells.Rows[1].GroupLevel > 0 || worksheet.Cells.Rows[2].GroupLevel > 0);
    }

    [Fact]
    public void GroupRows_WithCollapsed_ShouldGroupAndCollapse()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_rows_collapsed.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_rows_collapsed_output.xlsx");
        var result = _tool.Execute(
            "group_rows",
            workbookPath,
            startRow: 1,
            endRow: 3,
            isCollapsed: true,
            outputPath: outputPath);
        Assert.Contains("Rows 1-3 grouped", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void UngroupRows_ShouldUngroupRows()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_rows.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells.GroupRows(1, 3, false);
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_ungroup_rows_output.xlsx");
        var result = _tool.Execute(
            "ungroup_rows",
            workbookPath,
            startRow: 1,
            endRow: 3,
            outputPath: outputPath);
        Assert.Contains("Rows 1-3 ungrouped", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GroupColumns_ShouldGroupColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_columns.xlsx", 5, 10);
        var outputPath = CreateTestFilePath("test_group_columns_output.xlsx");
        var result = _tool.Execute(
            "group_columns",
            workbookPath,
            startColumn: 1,
            endColumn: 3,
            outputPath: outputPath);
        Assert.Contains("Columns 1-3 grouped", result);
        Assert.True(File.Exists(outputPath));

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Cells.Columns[1].GroupLevel > 0 || worksheet.Cells.Columns[2].GroupLevel > 0);
    }

    [Fact]
    public void UngroupColumns_ShouldUngroupColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_columns.xlsx", 5, 10);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells.GroupColumns(1, 3, false);
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_ungroup_columns_output.xlsx");
        var result = _tool.Execute(
            "ungroup_columns",
            workbookPath,
            startColumn: 1,
            endColumn: 3,
            outputPath: outputPath);
        Assert.Contains("Columns 1-3 ungrouped", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GroupRows_WithSheetIndex_ShouldGroupCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_sheet_index.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells[0, 0].Value = "Test";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_group_sheet_index_output.xlsx");
        var result = _tool.Execute(
            "group_rows",
            workbookPath,
            sheetIndex: 1,
            startRow: 0,
            endRow: 2,
            outputPath: outputPath);
        Assert.Contains("sheet 1", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GroupRows_SingleRow_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_single_row.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_single_row_output.xlsx");
        var result = _tool.Execute(
            "group_rows",
            workbookPath,
            startRow: 2,
            endRow: 2,
            outputPath: outputPath);
        Assert.Contains("Rows 2-2 grouped", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_invalid_op.xlsx", 5, 5);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid",
            workbookPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void GroupRows_MissingStartRow_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_missing_start.xlsx", 10, 5);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "group_rows",
            workbookPath,
            endRow: 3));
        Assert.Contains("requires parameter 'startRow'", exception.Message);
    }

    [Fact]
    public void GroupRows_MissingEndRow_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_missing_end.xlsx", 10, 5);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "group_rows",
            workbookPath,
            startRow: 1));
        Assert.Contains("requires parameter 'endRow'", exception.Message);
    }

    [Fact]
    public void GroupRows_StartGreaterThanEnd_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_invalid_range.xlsx", 10, 5);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "group_rows",
            workbookPath,
            startRow: 5,
            endRow: 2));
        Assert.Contains("cannot be greater than", exception.Message);
    }

    [Fact]
    public void GroupRows_NegativeStart_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_negative.xlsx", 10, 5);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "group_rows",
            workbookPath,
            startRow: -1,
            endRow: 3));
        Assert.Contains("cannot be negative", exception.Message);
    }

    [Fact]
    public void GroupColumns_MissingStartColumn_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_col_missing.xlsx", 5, 10);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "group_columns",
            workbookPath,
            endColumn: 3));
        Assert.Contains("requires parameter 'startColumn'", exception.Message);
    }

    [Fact]
    public void GroupColumns_StartGreaterThanEnd_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_col_invalid.xlsx", 5, 10);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "group_columns",
            workbookPath,
            startColumn: 5,
            endColumn: 2));
        Assert.Contains("cannot be greater than", exception.Message);
    }

    [Fact]
    public void GroupRows_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_invalid_sheet.xlsx", 10, 5);
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "group_rows",
            workbookPath,
            sheetIndex: 99,
            startRow: 0,
            endRow: 2));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GroupRows_WithSessionId_ShouldGroupInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_group_rows.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "group_rows",
            sessionId: sessionId,
            startRow: 1,
            endRow: 3);
        Assert.Contains("Rows 1-3 grouped", result);

        // Verify in-memory workbook has grouped rows
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.Rows[1].GroupLevel > 0 ||
                    workbook.Worksheets[0].Cells.Rows[2].GroupLevel > 0);
    }

    [Fact]
    public void UngroupRows_WithSessionId_ShouldUngroupInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_ungroup_rows.xlsx", 10, 5);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells.GroupRows(1, 3, false);
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "ungroup_rows",
            sessionId: sessionId,
            startRow: 1,
            endRow: 3);
        Assert.Contains("Rows 1-3 ungrouped", result);

        // Verify in-memory workbook has ungrouped rows
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(0, sessionWorkbook.Worksheets[0].Cells.Rows[1].GroupLevel);
    }

    [Fact]
    public void GroupColumns_WithSessionId_ShouldGroupInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_group_cols.xlsx", 5, 10);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "group_columns",
            sessionId: sessionId,
            startColumn: 1,
            endColumn: 3);
        Assert.Contains("Columns 1-3 grouped", result);

        // Verify in-memory workbook has grouped columns
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.Columns[1].GroupLevel > 0 ||
                    workbook.Worksheets[0].Cells.Columns[2].GroupLevel > 0);
    }

    [Fact]
    public void GroupRows_WithSessionId_ShouldNotModifyOriginalFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_group_original.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        _tool.Execute(
            "group_rows",
            sessionId: sessionId,
            startRow: 1,
            endRow: 3);

        // Assert - original file should not have grouped rows
        using var originalWorkbook = new Workbook(workbookPath);
        Assert.Equal(0, originalWorkbook.Worksheets[0].Cells.Rows[1].GroupLevel);
        Assert.Equal(0, originalWorkbook.Worksheets[0].Cells.Rows[2].GroupLevel);

        // But session workbook should have grouped rows
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(sessionWorkbook.Worksheets[0].Cells.Rows[1].GroupLevel > 0 ||
                    sessionWorkbook.Worksheets[0].Cells.Rows[2].GroupLevel > 0);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("group_rows", sessionId: "invalid_session_id", startRow: 1, endRow: 3));
    }

    #endregion
}