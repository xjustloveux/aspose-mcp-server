using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelGroupTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

    [Fact]
    public void GroupRows_ShouldGroupRows()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_group_rows_output.xlsx");
        var result = _tool.Execute("group_rows", workbookPath, startRow: 1, endRow: 3, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Rows 1-3 grouped", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.Rows[1].GroupLevel > 0 ||
                    workbook.Worksheets[0].Cells.Rows[2].GroupLevel > 0);
    }

    [Fact]
    public void UngroupRows_ShouldUngroupRows()
    {
        var workbookPath = CreateWorkbookWithGroupedRows("test_ungroup_rows.xlsx", 1, 3);
        var outputPath = CreateTestFilePath("test_ungroup_rows_output.xlsx");
        var result = _tool.Execute("ungroup_rows", workbookPath, startRow: 1, endRow: 3, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Rows 1-3 ungrouped", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(0, workbook.Worksheets[0].Cells.Rows[1].GroupLevel);
    }

    [Fact]
    public void GroupColumns_ShouldGroupColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_group_columns.xlsx", 5, 10);
        var outputPath = CreateTestFilePath("test_group_columns_output.xlsx");
        var result = _tool.Execute("group_columns", workbookPath, startColumn: 1, endColumn: 3, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Columns 1-3 grouped", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.Columns[1].GroupLevel > 0 ||
                    workbook.Worksheets[0].Cells.Columns[2].GroupLevel > 0);
    }

    [Fact]
    public void UngroupColumns_ShouldUngroupColumns()
    {
        var workbookPath = CreateWorkbookWithGroupedColumns("test_ungroup_columns.xlsx", 1, 3);
        var outputPath = CreateTestFilePath("test_ungroup_columns_output.xlsx");
        var result = _tool.Execute("ungroup_columns", workbookPath,
            startColumn: 1, endColumn: 3, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Columns 1-3 ungrouped", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(0, workbook.Worksheets[0].Cells.Columns[1].GroupLevel);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GROUP_ROWS")]
    [InlineData("Group_Rows")]
    [InlineData("group_rows")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation}.xlsx", 10, 5);
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, startRow: 1, endRow: 2, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("grouped", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx", 5, 5);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void GroupRows_WithSessionId_ShouldGroupInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_group_rows.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("group_rows", sessionId: sessionId, startRow: 1, endRow: 3);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Rows 1-3 grouped", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Rows 1-3 ungrouped", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(0, workbook.Worksheets[0].Cells.Rows[1].GroupLevel);
    }

    [Fact]
    public void GroupColumns_WithSessionId_ShouldGroupInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_group_cols.xlsx", 5, 10);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("group_columns", sessionId: sessionId, startColumn: 1, endColumn: 3);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Columns 1-3 grouped", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Rows 1-3 grouped", data.Message);
    }

    #endregion
}
