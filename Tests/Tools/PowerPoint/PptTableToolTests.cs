using Aspose.Slides;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptTableTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptTableToolTests : PptTestBase
{
    private readonly PptTableTool _tool;

    public PptTableToolTests()
    {
        _tool = new PptTableTool(SessionManager);
    }

    private static int FindTableShapeIndex(string pptPath, int slideIndex)
    {
        using var presentation = new Presentation(pptPath);
        var slide = presentation.Slides[slideIndex];
        var tableShape = slide.Shapes.OfType<ITable>().FirstOrDefault();
        return tableShape != null ? slide.Shapes.IndexOf(tableShape) : -1;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddTable_ShouldAddTableToSlide()
    {
        var pptPath = CreatePresentation("test_add_table.pptx");
        var outputPath = CreateTestFilePath("test_add_table_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, rows: 3, columns: 3, outputPath: outputPath);
        Assert.StartsWith("Table added to slide", result);
        Assert.Contains("3 rows", result);
        using var presentation = new Presentation(outputPath);
        var tables = presentation.Slides[0].Shapes.OfType<ITable>().ToList();
        Assert.NotEmpty(tables);
        Assert.Equal(3, tables[0].Rows.Count);
    }

    [Fact]
    public void DeleteTable_ShouldDeleteTable()
    {
        var pptPath = CreatePresentationWithTable("test_delete_table.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var outputPath = CreateTestFilePath("test_delete_table_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: shapeIndex, outputPath: outputPath);
        Assert.StartsWith("Table deleted from slide", result);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Slides[0].Shapes.OfType<ITable>());
    }

    [SkippableFact]
    public void GetContent_ShouldReturnTableContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");
        var pptPath = CreatePresentationWithTable("test_get_content.pptx");
        var shapeIndex = FindTableShapeIndex(pptPath, 0);
        var result = _tool.Execute("get_content", pptPath, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.Contains("R0C0", result);
        Assert.Contains("\"rowCount\"", result);
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
        using var presentation = new Presentation(outputPath);
        var table = presentation.Slides[0].Shapes.OfType<ITable>().First();
        Assert.Equal(3, table.Columns.Count);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, rows: 2, columns: 2, outputPath: outputPath);
        Assert.StartsWith("Table added to slide", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddTable_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialTableCount = ppt.Slides[0].Shapes.OfType<ITable>().Count();

        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, rows: 2, columns: 2);
        Assert.StartsWith("Table added to slide", result);
        Assert.Contains("session", result);

        var tablesAfter = ppt.Slides[0].Shapes.OfType<ITable>().Count();
        Assert.True(tablesAfter > initialTableCount);
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
        Assert.StartsWith("Table deleted from slide", result);
        Assert.Contains("session", result);
        Assert.Empty(ppt.Slides[0].Shapes.OfType<ITable>());
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
        var pptPath2 = CreatePresentation("test_session_table.pptx");

        var sessionId = OpenSession(pptPath2);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<ITable>().Count();

        var result = _tool.Execute("add", pptPath1, sessionId, slideIndex: 0, rows: 2, columns: 2);
        Assert.Contains("session", result);
        Assert.True(ppt.Slides[0].Shapes.OfType<ITable>().Count() > initialCount);
    }

    #endregion
}
