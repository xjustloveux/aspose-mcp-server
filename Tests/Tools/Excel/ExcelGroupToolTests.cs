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

    private string CreateWorkbookWithGroupedRows(string fileName, int startRow, int endRow)
    {
        var workbookPath = CreateExcelWorkbookWithData(fileName, 10, 5);
        using var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells.GroupRows(startRow, endRow, false);
        workbook.Save(workbookPath);
        return workbookPath;
    }

    private string CreateWorkbookWithGroupedColumns(string fileName, int startCol, int endCol)
    {
        var workbookPath = CreateExcelWorkbookWithData(fileName, 5, 10);
        using var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells.GroupColumns(startCol, endCol, false);
        workbook.Save(workbookPath);
        return workbookPath;
    }

    #region General

    [Fact]
    public void GroupRows_ShouldGroupRows()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_rows_output.xlsx");
        var result = _tool.Execute("group_rows", workbookPath, startRow: 1, endRow: 3, outputPath: outputPath);
        Assert.Contains("Rows 1-3 grouped", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.Rows[1].GroupLevel > 0 ||
                    workbook.Worksheets[0].Cells.Rows[2].GroupLevel > 0);
    }

    [Fact]
    public void GroupRows_WithCollapsed_ShouldGroupAndCollapse()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_rows_collapsed.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_rows_collapsed_output.xlsx");
        var result = _tool.Execute("group_rows", workbookPath, startRow: 1, endRow: 3,
            isCollapsed: true, outputPath: outputPath);
        Assert.Contains("Rows 1-3 grouped", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GroupRows_SingleRow_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_single_row.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_single_row_output.xlsx");
        var result = _tool.Execute("group_rows", workbookPath, startRow: 2, endRow: 2, outputPath: outputPath);
        Assert.Contains("Rows 2-2 grouped", result);
    }

    [Fact]
    public void GroupRows_WithSheetIndex_ShouldGroupCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_sheet.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells[0, 0].Value = "Test";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_group_sheet_output.xlsx");
        var result = _tool.Execute("group_rows", workbookPath, sheetIndex: 1,
            startRow: 0, endRow: 2, outputPath: outputPath);
        Assert.Contains("sheet 1", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(0, workbook.Worksheets[0].Cells.Rows[0].GroupLevel);
    }

    [Fact]
    public void UngroupRows_ShouldUngroupRows()
    {
        var workbookPath = CreateWorkbookWithGroupedRows("test_ungroup_rows.xlsx", 1, 3);
        var outputPath = CreateTestFilePath("test_ungroup_rows_output.xlsx");
        var result = _tool.Execute("ungroup_rows", workbookPath, startRow: 1, endRow: 3, outputPath: outputPath);
        Assert.Contains("Rows 1-3 ungrouped", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(0, workbook.Worksheets[0].Cells.Rows[1].GroupLevel);
    }

    [Fact]
    public void GroupColumns_ShouldGroupColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_columns.xlsx", 5, 10);
        var outputPath = CreateTestFilePath("test_group_columns_output.xlsx");
        var result = _tool.Execute("group_columns", workbookPath, startColumn: 1, endColumn: 3, outputPath: outputPath);
        Assert.Contains("Columns 1-3 grouped", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.Columns[1].GroupLevel > 0 ||
                    workbook.Worksheets[0].Cells.Columns[2].GroupLevel > 0);
    }

    [Fact]
    public void GroupColumns_WithCollapsed_ShouldGroupAndCollapse()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_columns_collapsed.xlsx", 5, 10);
        var outputPath = CreateTestFilePath("test_group_columns_collapsed_output.xlsx");
        var result = _tool.Execute("group_columns", workbookPath, startColumn: 1, endColumn: 3,
            isCollapsed: true, outputPath: outputPath);
        Assert.Contains("Columns 1-3 grouped", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GroupColumns_SingleColumn_ShouldSucceed()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_single_col.xlsx", 5, 10);
        var outputPath = CreateTestFilePath("test_group_single_col_output.xlsx");
        var result = _tool.Execute("group_columns", workbookPath, startColumn: 2, endColumn: 2, outputPath: outputPath);
        Assert.Contains("Columns 2-2 grouped", result);
    }

    [Fact]
    public void UngroupColumns_ShouldUngroupColumns()
    {
        var workbookPath = CreateWorkbookWithGroupedColumns("test_ungroup_columns.xlsx", 1, 3);
        var outputPath = CreateTestFilePath("test_ungroup_columns_output.xlsx");
        var result = _tool.Execute("ungroup_columns", workbookPath,
            startColumn: 1, endColumn: 3, outputPath: outputPath);
        Assert.Contains("Columns 1-3 ungrouped", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(0, workbook.Worksheets[0].Cells.Columns[1].GroupLevel);
    }

    [Theory]
    [InlineData("GROUP_ROWS")]
    [InlineData("Group_Rows")]
    [InlineData("group_rows")]
    public void Operation_ShouldBeCaseInsensitive_GroupRows(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 10, 5);
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, startRow: 1, endRow: 2, outputPath: outputPath);
        Assert.Contains("grouped", result);
    }

    [Theory]
    [InlineData("GROUP_COLUMNS")]
    [InlineData("Group_Columns")]
    [InlineData("group_columns")]
    public void Operation_ShouldBeCaseInsensitive_GroupColumns(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_col_{operation}.xlsx", 5, 10);
        var outputPath = CreateTestFilePath($"test_case_col_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, startColumn: 1, endColumn: 2, outputPath: outputPath);
        Assert.Contains("grouped", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx", 5, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void GroupRows_WithMissingStartRow_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_missing_start.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("group_rows", workbookPath, endRow: 3));
        Assert.Contains("requires parameter 'startRow'", ex.Message);
    }

    [Fact]
    public void GroupRows_WithMissingEndRow_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_missing_end.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("group_rows", workbookPath, startRow: 1));
        Assert.Contains("requires parameter 'endRow'", ex.Message);
    }

    [Fact]
    public void GroupRows_WithStartGreaterThanEnd_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_invalid_range.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("group_rows", workbookPath, startRow: 5, endRow: 2));
        Assert.Contains("cannot be greater than", ex.Message);
    }

    [Fact]
    public void GroupRows_WithNegativeStart_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_negative.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("group_rows", workbookPath, startRow: -1, endRow: 3));
        Assert.Contains("cannot be negative", ex.Message);
    }

    [Fact]
    public void GroupRows_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_invalid_sheet.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("group_rows", workbookPath, sheetIndex: 99, startRow: 0, endRow: 2));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void UngroupRows_WithMissingStartRow_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_missing_start.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("ungroup_rows", workbookPath, endRow: 3));
        Assert.Contains("requires parameter 'startRow'", ex.Message);
    }

    [Fact]
    public void UngroupRows_WithMissingEndRow_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_missing_end.xlsx", 10, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("ungroup_rows", workbookPath, startRow: 1));
        Assert.Contains("requires parameter 'endRow'", ex.Message);
    }

    [Fact]
    public void GroupColumns_WithMissingStartColumn_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_col_missing.xlsx", 5, 10);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("group_columns", workbookPath, endColumn: 3));
        Assert.Contains("requires parameter 'startColumn'", ex.Message);
    }

    [Fact]
    public void GroupColumns_WithMissingEndColumn_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_col_missing_end.xlsx", 5, 10);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("group_columns", workbookPath, startColumn: 1));
        Assert.Contains("requires parameter 'endColumn'", ex.Message);
    }

    [Fact]
    public void GroupColumns_WithStartGreaterThanEnd_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_col_invalid.xlsx", 5, 10);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("group_columns", workbookPath, startColumn: 5, endColumn: 2));
        Assert.Contains("cannot be greater than", ex.Message);
    }

    [Fact]
    public void GroupColumns_WithNegativeStart_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_col_negative.xlsx", 5, 10);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("group_columns", workbookPath, startColumn: -1, endColumn: 3));
        Assert.Contains("cannot be negative", ex.Message);
    }

    [Fact]
    public void UngroupColumns_WithMissingStartColumn_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_col_missing.xlsx", 5, 10);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("ungroup_columns", workbookPath, endColumn: 3));
        Assert.Contains("requires parameter 'startColumn'", ex.Message);
    }

    [Fact]
    public void UngroupColumns_WithMissingEndColumn_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_ungroup_col_missing_end.xlsx", 5, 10);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("ungroup_columns", workbookPath, startColumn: 1));
        Assert.Contains("requires parameter 'endColumn'", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("group_rows", "", startRow: 1, endRow: 3));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("group_rows", startRow: 1, endRow: 3));
    }

    #endregion

    #region Session

    [Fact]
    public void GroupRows_WithSessionId_ShouldGroupInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_group_rows.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("group_rows", sessionId: sessionId, startRow: 1, endRow: 3);
        Assert.Contains("Rows 1-3 grouped", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.Rows[1].GroupLevel > 0 ||
                    workbook.Worksheets[0].Cells.Rows[2].GroupLevel > 0);
    }

    [Fact]
    public void UngroupRows_WithSessionId_ShouldUngroupInMemory()
    {
        var workbookPath = CreateWorkbookWithGroupedRows("test_session_ungroup_rows.xlsx", 1, 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("ungroup_rows", sessionId: sessionId, startRow: 1, endRow: 3);
        Assert.Contains("Rows 1-3 ungrouped", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(0, workbook.Worksheets[0].Cells.Rows[1].GroupLevel);
    }

    [Fact]
    public void GroupColumns_WithSessionId_ShouldGroupInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_group_cols.xlsx", 5, 10);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("group_columns", sessionId: sessionId, startColumn: 1, endColumn: 3);
        Assert.Contains("Columns 1-3 grouped", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.Columns[1].GroupLevel > 0 ||
                    workbook.Worksheets[0].Cells.Columns[2].GroupLevel > 0);
    }

    [Fact]
    public void UngroupColumns_WithSessionId_ShouldUngroupInMemory()
    {
        var workbookPath = CreateWorkbookWithGroupedColumns("test_session_ungroup_cols.xlsx", 1, 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("ungroup_columns", sessionId: sessionId, startColumn: 1, endColumn: 3);
        Assert.Contains("Columns 1-3 ungrouped", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(0, workbook.Worksheets[0].Cells.Columns[1].GroupLevel);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("group_rows", sessionId: "invalid_session", startRow: 1, endRow: 3));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbookWithData("test_path_file.xlsx", 10, 5);
        var workbookPath2 = CreateExcelWorkbookWithData("test_session_file.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath2))
        {
            wb.Worksheets[0].Name = "SessionSheet";
            wb.Save(workbookPath2);
        }

        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("group_rows", workbookPath1, sessionId, startRow: 1, endRow: 3);
        Assert.Contains("session", result);
    }

    #endregion
}