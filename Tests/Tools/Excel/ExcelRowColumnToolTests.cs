using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelRowColumnTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelRowColumnToolTests : ExcelTestBase
{
    private readonly ExcelRowColumnTool _tool;

    public ExcelRowColumnToolTests()
    {
        _tool = new ExcelRowColumnTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void InsertRow_ShouldInsertRow()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_row.xlsx", 3);
        var outputPath = CreateTestFilePath("test_insert_row_output.xlsx");
        var result = _tool.Execute("insert_row", workbookPath, rowIndex: 1, count: 1, outputPath: outputPath);
        Assert.Contains("Inserted 1 row(s)", result);
    }

    [Fact]
    public void DeleteRow_ShouldDeleteRow()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_row.xlsx", 3);
        var outputPath = CreateTestFilePath("test_delete_row_output.xlsx");
        var result = _tool.Execute("delete_row", workbookPath, rowIndex: 1, outputPath: outputPath);
        Assert.Contains("Deleted 1 row(s)", result);
    }

    [Fact]
    public void InsertColumn_ShouldInsertColumn()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_column.xlsx", 3);
        var outputPath = CreateTestFilePath("test_insert_column_output.xlsx");
        var result = _tool.Execute("insert_column", workbookPath, columnIndex: 1, count: 1, outputPath: outputPath);
        Assert.Contains("Inserted 1 column(s)", result);
    }

    [Fact]
    public void DeleteColumn_ShouldDeleteColumn()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_column.xlsx", 3);
        var outputPath = CreateTestFilePath("test_delete_column_output.xlsx");
        var result = _tool.Execute("delete_column", workbookPath, columnIndex: 1, outputPath: outputPath);
        Assert.Contains("Deleted 1 column(s)", result);
    }

    [Fact]
    public void InsertCells_WithShiftDown_ShouldShiftDown()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_cells.xlsx", 3);
        var outputPath = CreateTestFilePath("test_insert_cells_output.xlsx");
        var result = _tool.Execute("insert_cells", workbookPath, range: "A1:B1", shiftDirection: "Down",
            outputPath: outputPath);
        Assert.Contains("inserted", result);
    }

    [Fact]
    public void DeleteCells_WithShiftUp_ShouldShiftUp()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_cells.xlsx", 3);
        var outputPath = CreateTestFilePath("test_delete_cells_output.xlsx");
        var result = _tool.Execute("delete_cells", workbookPath, range: "A1:B1", shiftDirection: "Up",
            outputPath: outputPath);
        Assert.Contains("deleted", result);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("INSERT_ROW")]
    [InlineData("Insert_Row")]
    [InlineData("insert_row")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation.Replace("_", "")}.xlsx", 3);
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, rowIndex: 1, outputPath: outputPath);
        Assert.Contains("Inserted 1 row(s)", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void InsertRow_WithSessionId_ShouldInsertInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_insert_row.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("insert_row", sessionId: sessionId, rowIndex: 1, count: 1);
        Assert.Contains("Inserted 1 row(s)", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void DeleteRow_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_delete_row.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete_row", sessionId: sessionId, rowIndex: 1, count: 1);
        Assert.Contains("Deleted 1 row(s)", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void InsertColumn_WithSessionId_ShouldInsertInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_insert_col.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("insert_column", sessionId: sessionId, columnIndex: 1, count: 1);
        Assert.Contains("Inserted 1 column(s)", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("insert_row", sessionId: "invalid_session", rowIndex: 1));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbookWithData("test_path_file.xlsx", 2);
        var workbookPath2 = CreateExcelWorkbookWithData("test_session_file.xlsx");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("insert_row", workbookPath1, sessionId, rowIndex: 0, count: 1);
        Assert.Contains("session", result);
    }

    #endregion
}
