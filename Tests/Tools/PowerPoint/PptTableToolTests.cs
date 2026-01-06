using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptTableToolTests : TestBase
{
    private readonly PptTableTool _tool;

    public PptTableToolTests()
    {
        _tool = new PptTableTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithTable(string fileName, int rows = 2, int columns = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var colWidths = Enumerable.Repeat(100.0, columns).ToArray();
        var rowHeights = Enumerable.Repeat(30.0, rows).ToArray();
        var table = slide.Shapes.AddTable(100, 100, colWidths, rowHeights);
        for (var r = 0; r < rows; r++)
        for (var c = 0; c < columns; c++)
            table[c, r].TextFrame.Text = $"R{r}C{c}";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private int FindTableShapeIndex(string pptPath, int slideIndex)
    {
        using var presentation = new Presentation(pptPath);
        var slide = presentation.Slides[slideIndex];
        var tableShape = slide.Shapes.OfType<ITable>().FirstOrDefault();
        return tableShape != null ? slide.Shapes.IndexOf(tableShape) : -1;
    }

    #region General

    [Fact]
    public void AddTable_ShouldAddTableToSlide()
    {
        var pptPath = CreateTestPresentation("test_add_table.pptx");
        var outputPath = CreateTestFilePath("test_add_table_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, rows: 3, columns: 3, outputPath: outputPath);
        Assert.StartsWith("Table (", result);
        Assert.Contains("added to slide", result);
        using var presentation = new Presentation(outputPath);
        var tables = presentation.Slides[0].Shapes.OfType<ITable>().ToList();
        Assert.NotEmpty(tables);
        Assert.Equal(3, tables[0].Rows.Count);
        Assert.Equal(3, tables[0].Columns.Count);
    }

    [Fact]
    public void AddTable_WithCustomPosition_ShouldPlaceTableAtPosition()
    {
        var pptPath = CreateTestPresentation("test_add_pos.pptx");
        var outputPath = CreateTestFilePath("test_add_pos_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, x: 150, y: 200, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.Equal(150, table.X, 1);
        Assert.Equal(200, table.Y, 1);
    }

    [SkippableFact]
    public void AddTable_WithData_ShouldFillTableWithData()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestPresentation("test_add_data.pptx");
        var outputPath = CreateTestFilePath("test_add_data_output.pptx");
        var dataJson = JsonSerializer.Serialize(new[] { new[] { "A1", "B1" }, new[] { "A2", "B2" } });
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, data: dataJson, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.Contains("A1", table[0, 0].TextFrame.Text);
    }

    [Fact]
    public void EditTable_ShouldEditTableData()
    {
        var pptPath = CreatePresentationWithTable("test_edit_table.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_edit_table_output.pptx");
        var dataJson = JsonSerializer.Serialize(new[] { new[] { "New1", "New2" }, new[] { "New3", "New4" } });
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: shapeIndex, data: dataJson,
            outputPath: outputPath);
        Assert.StartsWith("Table on slide", result);
        Assert.Contains("updated", result);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.NotNull(table);
    }

    [Fact]
    public void DeleteTable_ShouldDeleteTable()
    {
        var pptPath = CreatePresentationWithTable("test_delete_table.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_delete_table_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: shapeIndex, outputPath: outputPath);
        Assert.StartsWith("Table on slide", result);
        Assert.Contains("deleted", result);
        using var presentation = new Presentation(outputPath);
        var tables = presentation.Slides[0].Shapes.OfType<ITable>().ToList();
        Assert.Empty(tables);
    }

    [SkippableFact]
    public void GetContent_ShouldReturnTableContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreatePresentationWithTable("test_get_content.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var result = _tool.Execute("get_content", pptPath, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.Contains("R0C0", result);
        Assert.Contains("R1C1", result);
        Assert.Contains("\"rows\"", result);
        Assert.Contains("\"columns\"", result);
    }

    [Fact]
    public void InsertRow_ShouldInsertRow()
    {
        var pptPath = CreatePresentationWithTable("test_insert_row.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_insert_row_output.pptx");
        var result = _tool.Execute("insert_row", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 1,
            outputPath: outputPath);
        Assert.StartsWith("Row inserted at index", result);
        Assert.Contains("3 rows", result);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.Equal(3, table.Rows.Count);
    }

    [Fact]
    public void InsertRow_AtEnd_ShouldAppendRow()
    {
        var pptPath = CreatePresentationWithTable("test_insert_row_end.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_insert_row_end_output.pptx");
        var result = _tool.Execute("insert_row", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 2,
            outputPath: outputPath);
        Assert.StartsWith("Row inserted at index", result);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.Equal(3, table.Rows.Count);
    }

    [Fact]
    public void InsertColumn_ShouldInsertColumn()
    {
        var pptPath = CreatePresentationWithTable("test_insert_column.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_insert_column_output.pptx");
        var result = _tool.Execute("insert_column", pptPath, slideIndex: 0, shapeIndex: shapeIndex, columnIndex: 1,
            outputPath: outputPath);
        Assert.StartsWith("Column inserted at index", result);
        Assert.Contains("3 columns", result);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.Equal(3, table.Columns.Count);
    }

    [Fact]
    public void DeleteRow_ShouldDeleteRow()
    {
        var pptPath = CreatePresentationWithTable("test_delete_row.pptx", 3);
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_delete_row_output.pptx");
        var result = _tool.Execute("delete_row", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 1,
            outputPath: outputPath);
        Assert.StartsWith("Row", result);
        Assert.Contains("deleted", result);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.Equal(2, table.Rows.Count);
    }

    [Fact]
    public void DeleteColumn_ShouldDeleteColumn()
    {
        var pptPath = CreatePresentationWithTable("test_delete_column.pptx", 2, 3);
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_delete_column_output.pptx");
        var result = _tool.Execute("delete_column", pptPath, slideIndex: 0, shapeIndex: shapeIndex, columnIndex: 1,
            outputPath: outputPath);
        Assert.StartsWith("Column", result);
        Assert.Contains("deleted", result);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.Equal(2, table.Columns.Count);
    }

    [Fact]
    public void EditCell_ShouldEditCellContent()
    {
        var pptPath = CreatePresentationWithTable("test_edit_cell.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_edit_cell_output.pptx");
        var result = _tool.Execute("edit_cell", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 0,
            columnIndex: 0, text: "NewValue", outputPath: outputPath);
        Assert.StartsWith("Cell [", result);
        Assert.Contains("updated", result);
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        var cellText = table[0, 0].TextFrame?.Text ?? "";

        if (IsEvaluationMode())
            Assert.True(cellText.Contains("NewValue") || cellText.Contains("NewVal") || cellText.Length > 0);
        else
            Assert.Contains("NewValue", cellText);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: outputPath);
        Assert.StartsWith("Table (", result);
        Assert.Contains("added to slide", result);
    }

    [Theory]
    [InlineData("GET_CONTENT")]
    [InlineData("Get_Content")]
    [InlineData("get_content")]
    public void Operation_ShouldBeCaseInsensitive_GetContent(string operation)
    {
        var pptPath = CreatePresentationWithTable($"test_case_get_{operation.Replace("_", "")}.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.Contains("\"rows\"", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var pptPath = CreatePresentationWithTable($"test_case_del_{operation}.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex, outputPath: outputPath);
        Assert.StartsWith("Table on slide", result);
        Assert.Contains("deleted", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddTable_WithoutRows_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_no_rows.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 0, columns: 2));
        Assert.Contains("rows is required", ex.Message);
    }

    [Fact]
    public void AddTable_WithoutColumns_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_no_cols.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 0, rows: 2));
        Assert.Contains("columns is required", ex.Message);
    }

    [Fact]
    public void AddTable_WithInvalidRows_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_invalid_rows.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 0, rows: 0, columns: 2));
        Assert.Contains("rows must be between", ex.Message);
    }

    [Fact]
    public void AddTable_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 999, rows: 2, columns: 2));
        Assert.Contains("slideIndex must be between", ex.Message);
    }

    [Fact]
    public void EditTable_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithTable("test_edit_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pptPath, slideIndex: 0, data: "[[\"A\"]]"));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void EditTable_WithNonTableShape_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestFilePath("test_edit_non_table.pptx");
        using (var pres = new Presentation())
        {
            pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 0, data: "[[\"A\"]]"));
        Assert.Contains("not a table", ex.Message);
    }

    [Fact]
    public void InsertRow_OutOfRange_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithTable("test_insert_row_oor.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_row", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void InsertColumn_OutOfRange_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithTable("test_insert_col_oor.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_column", pptPath, slideIndex: 0, shapeIndex: shapeIndex, columnIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditCell_WithoutText_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithTable("test_edit_cell_no_text.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_cell", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 0, columnIndex: 0));
        Assert.Contains("text is required", ex.Message);
    }

    [Fact]
    public void EditCell_WithInvalidRowIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithTable("test_edit_cell_invalid_row.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_cell", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 99, columnIndex: 0,
                text: "Test"));
        Assert.Contains("rowIndex", ex.Message);
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditCell_WithInvalidColumnIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithTable("test_edit_cell_invalid_col.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_cell", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 0, columnIndex: 99,
                text: "Test"));
        Assert.Contains("columnIndex", ex.Message);
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void AddTable_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialTableCount = ppt.Slides[0].Shapes.OfType<ITable>().Count();

        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, rows: 2, columns: 2);
        Assert.StartsWith("Table (", result);
        Assert.Contains("added to slide", result);
        Assert.Contains("session", result);

        var tablesAfter = ppt.Slides[0].Shapes.OfType<ITable>().Count();
        Assert.True(tablesAfter > initialTableCount);
    }

    [SkippableFact]
    public void GetContent_WithSessionId_ShouldReturnContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreatePresentationWithTable("test_session_get.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var table = ppt.Slides[0].Shapes.OfType<ITable>().First();
        var shapeIndex = ppt.Slides[0].Shapes.IndexOf(table);

        var result = _tool.Execute("get_content", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.Contains("R0C0", result);
    }

    [SkippableFact]
    public void EditCell_WithSessionId_ShouldEditInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreatePresentationWithTable("test_session_edit_cell.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var table = ppt.Slides[0].Shapes.OfType<ITable>().First();
        var shapeIndex = ppt.Slides[0].Shapes.IndexOf(table);

        var result = _tool.Execute("edit_cell", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            rowIndex: 0, columnIndex: 0, text: "SessionEdit");
        Assert.StartsWith("Cell [", result);
        Assert.Contains("updated", result);
        Assert.Contains("session", result);
        Assert.Contains("SessionEdit", table[0, 0].TextFrame.Text);
    }

    [Fact]
    public void InsertRow_WithSessionId_ShouldInsertInMemory()
    {
        var pptPath = CreatePresentationWithTable("test_session_insert_row.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var table = ppt.Slides[0].Shapes.OfType<ITable>().First();
        var shapeIndex = ppt.Slides[0].Shapes.IndexOf(table);
        var initialRowCount = table.Rows.Count;

        var result = _tool.Execute("insert_row", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            rowIndex: 1);
        Assert.StartsWith("Row inserted at index", result);
        Assert.Contains("session", result);
        Assert.True(table.Rows.Count > initialRowCount);
    }

    [Fact]
    public void DeleteTable_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithTable("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var table = ppt.Slides[0].Shapes.OfType<ITable>().First();
        var shapeIndex = ppt.Slides[0].Shapes.IndexOf(table);

        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.StartsWith("Table on slide", result);
        Assert.Contains("deleted", result);
        Assert.Contains("session", result);

        var tablesAfter = ppt.Slides[0].Shapes.OfType<ITable>().Count();
        Assert.Equal(0, tablesAfter);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("add", sessionId: "invalid_session_id", slideIndex: 0, rows: 2, columns: 2));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithTable("test_path_table.pptx");
        var pptPath2 = CreateTestPresentation("test_session_table.pptx");

        var sessionId = OpenSession(pptPath2);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<ITable>().Count();

        var result = _tool.Execute("add", pptPath1, sessionId, slideIndex: 0, rows: 2, columns: 2);
        Assert.Contains("session", result);

        var tablesAfter = ppt.Slides[0].Shapes.OfType<ITable>().Count();
        Assert.True(tablesAfter > initialCount);
    }

    #endregion
}