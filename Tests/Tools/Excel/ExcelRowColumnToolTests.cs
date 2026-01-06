using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelRowColumnToolTests : ExcelTestBase
{
    private readonly ExcelRowColumnTool _tool;

    public ExcelRowColumnToolTests()
    {
        _tool = new ExcelRowColumnTool(SessionManager);
    }

    #region General

    [Fact]
    public void InsertRow_ShouldInsertRow()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_row.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalRowCount = originalWorkbook.Worksheets[0].Cells.Rows.Count;

        var outputPath = CreateTestFilePath("test_insert_row_output.xlsx");
        var result = _tool.Execute("insert_row", workbookPath, rowIndex: 1, count: 1, outputPath: outputPath);
        Assert.Contains("Inserted 1 row(s)", result);
        var workbook = new Workbook(outputPath);
        Assert.Equal(originalRowCount + 1, workbook.Worksheets[0].Cells.Rows.Count);
    }

    [Fact]
    public void InsertRow_WithMultiple_ShouldInsertMultipleRows()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_multiple_rows.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalRowCount = originalWorkbook.Worksheets[0].Cells.Rows.Count;

        var outputPath = CreateTestFilePath("test_insert_multiple_rows_output.xlsx");
        var result = _tool.Execute("insert_row", workbookPath, rowIndex: 1, count: 3, outputPath: outputPath);
        Assert.Contains("Inserted 3 row(s)", result);
        var workbook = new Workbook(outputPath);
        Assert.Equal(originalRowCount + 3, workbook.Worksheets[0].Cells.Rows.Count);
    }

    [Fact]
    public void InsertRow_WithSheetIndex_ShouldInsertInCorrectSheet()
    {
        var workbookPath = CreateTestFilePath("test_insert_row_sheet.xlsx");
        using (var wb = new Workbook())
        {
            wb.Worksheets[0].Cells["A1"].Value = "Sheet1";
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells["A1"].Value = "Sheet2-R0";
            wb.Worksheets[1].Cells["A2"].Value = "Sheet2-R1";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_insert_row_sheet_output.xlsx");
        _tool.Execute("insert_row", workbookPath, sheetIndex: 1, rowIndex: 1, outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        Assert.Equal("Sheet2-R0", workbook.Worksheets[1].Cells["A1"].Value?.ToString());
    }

    [Fact]
    public void DeleteRow_ShouldDeleteRow()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_row.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalRowCount = originalWorkbook.Worksheets[0].Cells.Rows.Count;

        var outputPath = CreateTestFilePath("test_delete_row_output.xlsx");
        var result = _tool.Execute("delete_row", workbookPath, rowIndex: 1, outputPath: outputPath);
        Assert.Contains("Deleted 1 row(s)", result);
        var workbook = new Workbook(outputPath);
        Assert.Equal(originalRowCount - 1, workbook.Worksheets[0].Cells.Rows.Count);
    }

    [Fact]
    public void DeleteRow_WithMultiple_ShouldDeleteMultipleRows()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_multiple_rows.xlsx");
        var originalWorkbook = new Workbook(workbookPath);
        var originalRowCount = originalWorkbook.Worksheets[0].Cells.Rows.Count;

        var outputPath = CreateTestFilePath("test_delete_multiple_rows_output.xlsx");
        var result = _tool.Execute("delete_row", workbookPath, rowIndex: 1, count: 2, outputPath: outputPath);
        Assert.Contains("Deleted 2 row(s)", result);
        var workbook = new Workbook(outputPath);
        Assert.Equal(originalRowCount - 2, workbook.Worksheets[0].Cells.Rows.Count);
    }

    [Fact]
    public void InsertColumn_ShouldInsertColumn()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_column.xlsx", 3);
        var outputPath = CreateTestFilePath("test_insert_column_output.xlsx");
        var result = _tool.Execute("insert_column", workbookPath, columnIndex: 1, count: 1, outputPath: outputPath);
        Assert.Contains("Inserted 1 column(s)", result);
        var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.MaxColumn >= 2);
    }

    [Fact]
    public void InsertColumn_WithMultiple_ShouldInsertMultipleColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_multiple_cols.xlsx", 3);
        var outputPath = CreateTestFilePath("test_insert_multiple_cols_output.xlsx");
        var result = _tool.Execute("insert_column", workbookPath, columnIndex: 1, count: 2, outputPath: outputPath);
        Assert.Contains("Inserted 2 column(s)", result);
        var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.MaxColumn >= 4);
    }

    [Fact]
    public void DeleteColumn_ShouldDeleteColumn()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_column.xlsx", 3);
        var outputPath = CreateTestFilePath("test_delete_column_output.xlsx");
        var result = _tool.Execute("delete_column", workbookPath, columnIndex: 1, outputPath: outputPath);
        Assert.Contains("Deleted 1 column(s)", result);
        var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.MaxColumn <= 2);
    }

    [Fact]
    public void DeleteColumn_WithMultiple_ShouldDeleteMultipleColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_multiple_cols.xlsx", 3, 5);
        var outputPath = CreateTestFilePath("test_delete_multiple_cols_output.xlsx");
        var result = _tool.Execute("delete_column", workbookPath, columnIndex: 1, count: 2, outputPath: outputPath);
        Assert.Contains("Deleted 2 column(s)", result);
        var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].Cells.MaxColumn <= 2);
    }

    [Fact]
    public void InsertCells_WithShiftDown_ShouldShiftDown()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_cells.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalA1 = originalWorkbook.Worksheets[0].Cells["A1"].Value;
        var originalB1 = originalWorkbook.Worksheets[0].Cells["B1"].Value;

        var outputPath = CreateTestFilePath("test_insert_cells_output.xlsx");
        var result = _tool.Execute("insert_cells", workbookPath, range: "A1:B1", shiftDirection: "Down",
            outputPath: outputPath);
        Assert.Contains("inserted", result);
        Assert.True(File.Exists(outputPath));

        var resultWorkbook = new Workbook(outputPath);
        var a1After = resultWorkbook.Worksheets[0].Cells["A1"].Value;
        var a2After = resultWorkbook.Worksheets[0].Cells["A2"].Value;
        var b2After = resultWorkbook.Worksheets[0].Cells["B2"].Value;
        Assert.True(a1After == null || a1After.ToString() == "", "A1 should be empty after insert");
        Assert.Equal(originalA1, a2After);
        Assert.Equal(originalB1, b2After);
    }

    [Fact]
    public void InsertCells_WithShiftRight_ShouldShiftRight()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_cells_right.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalA1 = originalWorkbook.Worksheets[0].Cells["A1"].Value;
        var originalA2 = originalWorkbook.Worksheets[0].Cells["A2"].Value;

        var outputPath = CreateTestFilePath("test_insert_cells_right_output.xlsx");
        var result = _tool.Execute("insert_cells", workbookPath, range: "A1:A2", shiftDirection: "Right",
            outputPath: outputPath);
        Assert.Contains("inserted", result);
        Assert.True(File.Exists(outputPath));

        var resultWorkbook = new Workbook(outputPath);
        var a1After = resultWorkbook.Worksheets[0].Cells["A1"].Value;
        var b1After = resultWorkbook.Worksheets[0].Cells["B1"].Value;
        var b2After = resultWorkbook.Worksheets[0].Cells["B2"].Value;
        Assert.True(a1After == null || a1After.ToString() == "", "A1 should be empty after insert");
        Assert.Equal(originalA1, b1After);
        Assert.Equal(originalA2, b2After);
    }

    [Fact]
    public void InsertCells_SingleCell_ShouldWork()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_single_cell.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalB2 = originalWorkbook.Worksheets[0].Cells["B2"].Value;

        var outputPath = CreateTestFilePath("test_insert_single_cell_output.xlsx");
        var result = _tool.Execute("insert_cells", workbookPath, range: "B2", shiftDirection: "Down",
            outputPath: outputPath);
        Assert.Contains("inserted", result);
        Assert.True(File.Exists(outputPath));

        var resultWorkbook = new Workbook(outputPath);
        var b2After = resultWorkbook.Worksheets[0].Cells["B2"].Value;
        var b3After = resultWorkbook.Worksheets[0].Cells["B3"].Value;
        Assert.True(b2After == null || b2After.ToString() == "", "B2 should be empty after insert");
        Assert.Equal(originalB2, b3After);
    }

    [Fact]
    public void DeleteCells_WithShiftUp_ShouldShiftUp()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_cells.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalA2 = originalWorkbook.Worksheets[0].Cells["A2"].Value;
        var originalB2 = originalWorkbook.Worksheets[0].Cells["B2"].Value;

        var outputPath = CreateTestFilePath("test_delete_cells_output.xlsx");
        var result = _tool.Execute("delete_cells", workbookPath, range: "A1:B1", shiftDirection: "Up",
            outputPath: outputPath);
        Assert.Contains("deleted", result);
        Assert.True(File.Exists(outputPath));

        var resultWorkbook = new Workbook(outputPath);
        var a1After = resultWorkbook.Worksheets[0].Cells["A1"].Value;
        var b1After = resultWorkbook.Worksheets[0].Cells["B1"].Value;
        Assert.Equal(originalA2, a1After);
        Assert.Equal(originalB2, b1After);
    }

    [Fact]
    public void DeleteCells_WithShiftLeft_ShouldShiftLeft()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_cells_left.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalC1 = originalWorkbook.Worksheets[0].Cells["C1"].Value;
        var originalC2 = originalWorkbook.Worksheets[0].Cells["C2"].Value;

        var outputPath = CreateTestFilePath("test_delete_cells_left_output.xlsx");
        var result = _tool.Execute("delete_cells", workbookPath, range: "B1:B2", shiftDirection: "Left",
            outputPath: outputPath);
        Assert.Contains("deleted", result);
        Assert.True(File.Exists(outputPath));

        var resultWorkbook = new Workbook(outputPath);
        var b1After = resultWorkbook.Worksheets[0].Cells["B1"].Value;
        var b2After = resultWorkbook.Worksheets[0].Cells["B2"].Value;
        Assert.Equal(originalC1, b1After);
        Assert.Equal(originalC2, b2After);
    }

    [Fact]
    public void DeleteCells_SingleCell_ShouldWork()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_single_cell.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var originalB3 = originalWorkbook.Worksheets[0].Cells["B3"].Value;

        var outputPath = CreateTestFilePath("test_delete_single_cell_output.xlsx");
        var result = _tool.Execute("delete_cells", workbookPath, range: "B2", shiftDirection: "Up",
            outputPath: outputPath);
        Assert.Contains("deleted", result);
        Assert.True(File.Exists(outputPath));

        var resultWorkbook = new Workbook(outputPath);
        var b2After = resultWorkbook.Worksheets[0].Cells["B2"].Value;
        Assert.Equal(originalB3, b2After);
    }

    [Theory]
    [InlineData("INSERT_ROW")]
    [InlineData("Insert_Row")]
    [InlineData("insert_row")]
    public void Operation_ShouldBeCaseInsensitive_InsertRow(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation.Replace("_", "")}.xlsx", 3);
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, rowIndex: 1, outputPath: outputPath);
        Assert.Contains("Inserted 1 row(s)", result);
    }

    [Theory]
    [InlineData("INSERT_CELLS")]
    [InlineData("Insert_Cells")]
    [InlineData("insert_cells")]
    public void Operation_ShouldBeCaseInsensitive_InsertCells(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation.Replace("_", "")}.xlsx", 3);
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1", shiftDirection: "Down",
            outputPath: outputPath);
        Assert.Contains("inserted", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_unknown_op.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void SetColumnWidth_ShouldThrowWithGuidance()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_column_width.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set_column_width", workbookPath));
        Assert.Contains("excel_view_settings", ex.Message);
    }

    [Fact]
    public void InsertCells_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_cells_no_range.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_cells", workbookPath, shiftDirection: "Down"));
        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void InsertCells_WithMissingShiftDirection_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_insert_cells_no_shift.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_cells", workbookPath, range: "A1:B2"));
        Assert.Contains("shiftDirection is required", ex.Message);
    }

    [Fact]
    public void DeleteCells_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_cells_no_range.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_cells", workbookPath, shiftDirection: "Up"));
        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void DeleteCells_WithMissingShiftDirection_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_cells_no_shift.xlsx", 3);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_cells", workbookPath, range: "A1:B2"));
        Assert.Contains("shiftDirection is required", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("insert_row", rowIndex: 1));
    }

    #endregion

    #region Session

    [Fact]
    public void InsertRow_WithSessionId_ShouldInsertInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_insert_row.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("insert_row", sessionId: sessionId, rowIndex: 1, count: 1);
        Assert.Contains("Inserted 1 row(s)", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.Rows.Count >= 4);
    }

    [Fact]
    public void DeleteRow_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_delete_row.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete_row", sessionId: sessionId, rowIndex: 1, count: 1);
        Assert.Contains("Deleted 1 row(s)", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.Rows.Count <= 4);
    }

    [Fact]
    public void InsertColumn_WithSessionId_ShouldInsertInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_insert_col.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("insert_column", sessionId: sessionId, columnIndex: 1, count: 1);
        Assert.Contains("Inserted 1 column(s)", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.MaxColumn >= 3);
    }

    [Fact]
    public void DeleteColumn_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_delete_col.xlsx", 3, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete_column", sessionId: sessionId, columnIndex: 1, count: 1);
        Assert.Contains("Deleted 1 column(s)", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells.MaxColumn <= 3);
    }

    [Fact]
    public void InsertCells_WithSessionId_ShouldInsertInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_insert_cells.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("insert_cells", sessionId: sessionId, range: "A1:B2", shiftDirection: "Down");
        Assert.Contains("inserted", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void DeleteCells_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_delete_cells.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete_cells", sessionId: sessionId, range: "A1:B2", shiftDirection: "Up");
        Assert.Contains("deleted", result);
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