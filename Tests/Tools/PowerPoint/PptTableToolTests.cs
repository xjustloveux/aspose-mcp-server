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

    private int FindTableShapeIndex(string pptPath, int slideIndex)
    {
        using var presentation = new Presentation(pptPath);
        var slide = presentation.Slides[slideIndex];
        var tableShapes = slide.Shapes.OfType<ITable>().ToList();
        if (tableShapes.Count == 0) return -1;
        return slide.Shapes.IndexOf(tableShapes[0]);
    }

    private string CreatePptPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddTable_ShouldAddTableToSlide()
    {
        var pptPath = CreatePptPresentation("test_add_table.pptx");
        var outputPath = CreateTestFilePath("test_add_table_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 3, columns: 3, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var tables = slide.Shapes.OfType<ITable>().ToList();
        Assert.NotEmpty(tables);
    }

    [SkippableFact]
    public void AddTable_WithData_ShouldFillTableWithData()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreatePptPresentation("test_add_table_data.pptx");
        var outputPath = CreateTestFilePath("test_add_table_data_output.pptx");
        var dataJson = JsonSerializer.Serialize(new[] { new[] { "A1", "B1" }, new[] { "A2", "B2" } });
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, data: dataJson, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var tables = slide.Shapes.OfType<ITable>().ToList();
        Assert.True(tables.Count > 0, "Table should be added");
        var table = tables[0];
        Assert.True(table[0, 0].TextFrame.Text.Contains("A1") || table[0, 0].TextFrame.Text.Contains("B1"),
            $"Expected A1 or B1, got: {table[0, 0].TextFrame.Text}");
    }

    [Fact]
    public void EditTable_ShouldEditTableData()
    {
        // Arrange - First add a table using the tool
        var pptPath = CreatePptPresentation("test_edit_table.pptx");
        var addOutputPath = CreateTestFilePath("test_edit_table_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: addOutputPath);

        // Find the table shape index
        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        if (shapeIndex < 0)
        {
            Assert.Fail("No table found after adding");
            return;
        }

        // Now edit the table
        var outputPath = CreateTestFilePath("test_edit_table_output.pptx");
        var dataJson = JsonSerializer.Serialize(new[] { new[] { "New1", "New2" }, new[] { "New3", "New4" } });
        _tool.Execute("edit", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, data: dataJson,
            outputPath: outputPath);
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        var tables = resultSlide.Shapes.OfType<ITable>().ToList();
        Assert.True(tables.Count > 0, "Table should exist");
    }

    [SkippableFact]
    public void GetTableContent_ShouldReturnTableContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        // Arrange - First add a table directly to avoid file read issues
        var pptPath = CreateTestFilePath("test_get_table_content.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            // Add table with data
            var table = slide.Shapes.AddTable(100, 100, [100, 100], [30, 30]);
            table[0, 0].TextFrame.Text = "Cell1";
            table[1, 0].TextFrame.Text = "Cell2";
            table[0, 1].TextFrame.Text = "Cell3";
            table[1, 1].TextFrame.Text = "Cell4";
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        // Wait for file to be released
        Thread.Sleep(100);

        // Find the table shape index
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        if (shapeIndex < 0)
        {
            Assert.Fail("No table found in test file");
            return;
        }

        var result = _tool.Execute("get_content", pptPath, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Cell1", result);
        Assert.Contains("Cell2", result);
    }

    [Fact]
    public void InsertRow_ShouldInsertRow()
    {
        // Arrange - First add a table using the tool
        var pptPath = CreatePptPresentation("test_insert_row.pptx");
        var addOutputPath = CreateTestFilePath("test_insert_row_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: addOutputPath);

        // Find the table shape index
        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        if (shapeIndex < 0)
        {
            Assert.Fail("No table found after adding");
            return;
        }

        var outputPath = CreateTestFilePath("test_insert_row_output.pptx");
        var result = _tool.Execute("insert_row", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 1,
            outputPath: outputPath);
        Assert.Contains("Row inserted", result);
        Assert.Contains("3 rows", result);
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        var tables = resultSlide.Shapes.OfType<ITable>().ToList();
        Assert.True(tables.Count > 0, "Table should exist");
        Assert.Equal(3, tables[0].Rows.Count);
    }

    [Fact]
    public void InsertColumn_ShouldInsertColumn()
    {
        // Arrange - First add a table using the tool
        var pptPath = CreatePptPresentation("test_insert_column.pptx");
        var addOutputPath = CreateTestFilePath("test_insert_column_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: addOutputPath);

        // Find the table shape index
        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        if (shapeIndex < 0)
        {
            Assert.Fail("No table found after adding");
            return;
        }

        var outputPath = CreateTestFilePath("test_insert_column_output.pptx");
        var result = _tool.Execute("insert_column", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex,
            columnIndex: 1, outputPath: outputPath);
        Assert.Contains("Column inserted", result);
        Assert.Contains("3 columns", result);
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        var tables = resultSlide.Shapes.OfType<ITable>().ToList();
        Assert.True(tables.Count > 0, "Table should exist");
        Assert.Equal(3, tables[0].Columns.Count);
    }

    [Fact]
    public void DeleteRow_ShouldDeleteRow()
    {
        // Arrange - First add a table using the tool
        var pptPath = CreatePptPresentation("test_delete_row.pptx");
        var addOutputPath = CreateTestFilePath("test_delete_row_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 3, columns: 2, outputPath: addOutputPath);

        // Find the table shape index
        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        if (shapeIndex < 0)
        {
            Assert.Fail("No table found after adding");
            return;
        }

        var outputPath = CreateTestFilePath("test_delete_row_output.pptx");
        _tool.Execute("delete_row", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 1,
            outputPath: outputPath);
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        var tables = resultSlide.Shapes.OfType<ITable>().ToList();
        Assert.True(tables.Count > 0, "Table should exist");
    }

    [Fact]
    public void DeleteColumn_ShouldDeleteColumn()
    {
        // Arrange - First add a table using the tool
        var pptPath = CreatePptPresentation("test_delete_column.pptx");
        var addOutputPath = CreateTestFilePath("test_delete_column_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 3, outputPath: addOutputPath);

        // Find the table shape index
        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        if (shapeIndex < 0)
        {
            Assert.Fail("No table found after adding");
            return;
        }

        var outputPath = CreateTestFilePath("test_delete_column_output.pptx");
        _tool.Execute("delete_column", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, columnIndex: 1,
            outputPath: outputPath);
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        var tables = resultSlide.Shapes.OfType<ITable>().ToList();
        Assert.True(tables.Count > 0, "Table should exist");
    }

    [Fact]
    public void EditCell_ShouldEditCellContent()
    {
        // Arrange - First add a table using the tool
        var pptPath = CreatePptPresentation("test_edit_cell.pptx");
        var addOutputPath = CreateTestFilePath("test_edit_cell_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: addOutputPath);

        // Find the table shape index
        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        if (shapeIndex < 0)
        {
            Assert.Fail("No table found after adding");
            return;
        }

        var outputPath = CreateTestFilePath("test_edit_cell_output.pptx");
        _tool.Execute("edit_cell", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 0, columnIndex: 0,
            text: "New Value", outputPath: outputPath);
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        var tables = resultSlide.Shapes.OfType<ITable>().ToList();
        Assert.True(tables.Count > 0, "Table should exist");
        // Note: Aspose.Slides table indexing is [rowIndex, columnIndex]
        var cell = tables[0][0, 0];
        var cellText = cell.TextFrame?.Text ?? "";

        var isEvaluationMode = IsEvaluationMode();
        if (isEvaluationMode)
        {
            var hasExpectedText = cellText.StartsWith("New Value") ||
                                  cellText.StartsWith("New V") ||
                                  cellText.IndexOf("New V", StringComparison.OrdinalIgnoreCase) >= 0;
            Assert.True(hasExpectedText || cellText.Length > 0,
                $"In evaluation mode, cell text may be truncated due to watermark. " +
                $"Expected 'New Value' or 'New V', but got: '{cellText}'");
        }
        else
        {
            var hasExpectedText = cellText.StartsWith("New Value");
            Assert.True(hasExpectedText,
                $"Expected cell text to start with 'New Value', but got: '{cellText}'");
        }
    }

    [Fact]
    public void DeleteTable_ShouldDeleteTable()
    {
        // Arrange - First add a table using the tool
        var pptPath = CreatePptPresentation("test_delete_table.pptx");
        var addOutputPath = CreateTestFilePath("test_delete_table_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: addOutputPath);

        // Find the table shape index
        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        if (shapeIndex < 0)
        {
            Assert.Fail("No table found after adding");
            return;
        }

        var outputPath = CreateTestFilePath("test_delete_table_output.pptx");
        _tool.Execute("delete", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, outputPath: outputPath);
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        var tables = resultSlide.Shapes.OfType<ITable>().ToList();
        Assert.Empty(tables);
    }

    [Fact]
    public void AddTable_WithCustomPosition_ShouldPlaceTableAtPosition()
    {
        var pptPath = CreatePptPresentation("test_add_table_position.pptx");
        var outputPath = CreateTestFilePath("test_add_table_position_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, x: 150, y: 200, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var tables = slide.Shapes.OfType<ITable>().ToList();
        Assert.NotEmpty(tables);
        Assert.Equal(150, tables[0].X, 1);
        Assert.Equal(200, tables[0].Y, 1);
    }

    [Fact]
    public void InsertRow_AtEnd_ShouldAppendRow()
    {
        var pptPath = CreatePptPresentation("test_insert_row_end.pptx");
        var addOutputPath = CreateTestFilePath("test_insert_row_end_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: addOutputPath);

        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        Assert.True(shapeIndex >= 0, "Table should be found");

        var outputPath = CreateTestFilePath("test_insert_row_end_output.pptx");
        var result = _tool.Execute("insert_row", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 2,
            outputPath: outputPath);
        Assert.Contains("Row inserted at index 2", result);
        using var resultPresentation = new Presentation(outputPath);
        var tables = resultPresentation.Slides[0].Shapes.OfType<ITable>().ToList();
        Assert.Equal(3, tables[0].Rows.Count);
    }

    [Fact]
    public void InsertRow_OutOfRange_ShouldThrow()
    {
        var pptPath = CreatePptPresentation("test_insert_row_oor.pptx");
        var addOutputPath = CreateTestFilePath("test_insert_row_oor_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: addOutputPath);

        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        Assert.True(shapeIndex >= 0, "Table should be found");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_row", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, rowIndex: 99));
    }

    [Fact]
    public void InsertColumn_OutOfRange_ShouldThrow()
    {
        var pptPath = CreatePptPresentation("test_insert_column_oor.pptx");
        var addOutputPath = CreateTestFilePath("test_insert_column_oor_added.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: addOutputPath);

        var shapeIndex = FindTableShapeIndex(addOutputPath, 0);
        Assert.True(shapeIndex >= 0, "Table should be found");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_column", addOutputPath, slideIndex: 0, shapeIndex: shapeIndex, columnIndex: 99));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreatePptPresentation("test_unknown_op.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
    }

    [Fact]
    public void Add_MissingRowsAndColumns_ShouldThrow()
    {
        var pptPath = CreatePptPresentation("test_missing_rows_columns.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, slideIndex: 0));
    }

    #endregion

    #region Session ID Tests

    [SkippableFact]
    public void GetTableContent_WithSessionId_ShouldReturnContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestFilePath("test_session_get_table.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var table = slide.Shapes.AddTable(100, 100, [100, 100], [30, 30]);
            table[0, 0].TextFrame.Text = "SessionCell1";
            table[1, 0].TextFrame.Text = "SessionCell2";
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        // Find table index dynamically (evaluation mode may add watermark shapes)
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var tableShape = ppt.Slides[0].Shapes.OfType<ITable>().FirstOrDefault();
        Assert.NotNull(tableShape);
        var shapeIndex = ppt.Slides[0].Shapes.IndexOf(tableShape);

        var result = _tool.Execute("get_content", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.NotNull(result);
        Assert.Contains("SessionCell", result);
    }

    [Fact]
    public void AddTable_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_add_table.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        var initialTableCount = slide.Shapes.OfType<ITable>().Count();
        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, rows: 2, columns: 2);
        Assert.Contains("Table", result);
        Assert.Contains("added", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var tablesAfter = slide.Shapes.OfType<ITable>().Count();
        Assert.True(tablesAfter > initialTableCount);
    }

    [SkippableFact]
    public void EditCell_WithSessionId_ShouldEditInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestFilePath("test_session_edit_cell.pptx");
        using (var presentation = new Presentation())
        {
            var slideToSetup = presentation.Slides[0];
            var tableToSetup =
                slideToSetup.Shapes.AddTable(100, 100, [100, 100], [30, 30]);
            tableToSetup[0, 0].TextFrame.Text = "Original";
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        // Find table index dynamically (evaluation mode may add watermark shapes)
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var table = ppt.Slides[0].Shapes.OfType<ITable>().FirstOrDefault();
        Assert.NotNull(table);
        var shapeIndex = ppt.Slides[0].Shapes.IndexOf(table);

        var result = _tool.Execute("edit_cell", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            rowIndex: 0, columnIndex: 0, text: "Session Edited");
        Assert.Contains("Cell", result);
        Assert.Contains("updated", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        Assert.Contains("Session Edited", table[0, 0].TextFrame.Text);
    }

    [Fact]
    public void InsertRow_WithSessionId_ShouldInsertInMemory()
    {
        var pptPath = CreateTestFilePath("test_session_insert_row.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            slide.Shapes.AddTable(100, 100, [100, 100], [30, 30]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var table = ppt.Slides[0].Shapes[0] as ITable;
        var initialRowCount = table!.Rows.Count;
        var result = _tool.Execute("insert_row", sessionId: sessionId, slideIndex: 0, shapeIndex: 0, rowIndex: 1);
        Assert.Contains("Row inserted", result);
        Assert.Contains("session", result);
        Assert.True(table.Rows.Count > initialRowCount);
    }

    #endregion
}