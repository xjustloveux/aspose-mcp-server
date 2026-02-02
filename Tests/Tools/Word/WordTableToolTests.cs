using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Table;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordTableTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordTableToolTests : WordTestBase
{
    private readonly WordTableTool _tool;

    public WordTableToolTests()
    {
        _tool = new WordTableTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Create_ShouldCreateTableAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_create_table.docx");
        var outputPath = CreateTestFilePath("test_create_table_output.docx");
        _tool.Execute("create", docPath, outputPath: outputPath, rows: 3, columns: 4);
        var doc = new Document(outputPath);
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Single(tables);
        Assert.Equal(3, tables[0].Rows.Count);
        Assert.Equal(4, tables[0].FirstRow.Cells.Count);
    }

    [Fact]
    public void Get_ShouldReturnTablesFromFile()
    {
        var docPath = CreateWordDocument("test_get_tables.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var result = _tool.Execute("get", docPath);
        var data = GetResultData<GetTablesWordResult>(result);
        Assert.Equal(1, data.Count);
    }

    [Fact]
    public void Delete_ShouldDeleteTableAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_delete_table.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_table_output.docx");
        _tool.Execute("delete", docPath, outputPath: outputPath, tableIndex: 0);
        var doc = new Document(outputPath);
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Empty(tables);
    }

    [Fact]
    public void InsertRow_ShouldInsertRowAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_insert_row.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var outputPath = CreateTestFilePath("test_insert_row_output.docx");
        _tool.Execute("insert_row", docPath, outputPath: outputPath, tableIndex: 0, rowIndex: 0);
        var doc = new Document(outputPath);
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Equal(2, tables[0].Rows.Count);
    }

    [Fact]
    public void DeleteRow_ShouldDeleteRowAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_delete_row.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Row 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Row 2");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_row_output.docx");
        _tool.Execute("delete_row", docPath, outputPath: outputPath, tableIndex: 0, rowIndex: 1);
        var doc = new Document(outputPath);
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Single(tables[0].Rows);
    }

    [Fact]
    public void InsertColumn_ShouldInsertColumnAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_insert_col.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var outputPath = CreateTestFilePath("test_insert_col_output.docx");
        _tool.Execute("insert_column", docPath, outputPath: outputPath, tableIndex: 0, columnIndex: 0);
        var doc = new Document(outputPath);
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Equal(2, tables[0].FirstRow.Cells.Count);
    }

    [Fact]
    public void DeleteColumn_ShouldDeleteColumnAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_delete_col.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Col 1");
        builder.InsertCell();
        builder.Write("Col 2");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_col_output.docx");
        _tool.Execute("delete_column", docPath, outputPath: outputPath, tableIndex: 0, columnIndex: 1);
        var doc = new Document(outputPath);
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Single(tables[0].FirstRow.Cells);
    }

    [Fact]
    public void GetStructure_ShouldReturnTableStructureFromFile()
    {
        var docPath = CreateWordDocument("test_get_structure.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var result = _tool.Execute("get_structure", docPath, tableIndex: 0);
        var data = GetResultData<GetTableStructureWordResult>(result);
        Assert.Contains("Table", data.Content);
        Assert.Contains("Rows", data.Content);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.docx");
        _tool.Execute(operation, docPath, outputPath: outputPath, rows: 2, columns: 2);
        var doc = new Document(outputPath);
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Single(tables);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_unknown_op.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldReturnTables()
    {
        var docPath = CreateWordDocument("test_session_get_tables.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Session Table");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get", sessionId: sessionId);

        var data = GetResultData<GetTablesWordResult>(result);
        Assert.Equal(1, data.Count);
        var output = GetResultOutput<GetTablesWordResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Create_WithSessionId_ShouldCreateTableInMemory()
    {
        var docPath = CreateWordDocument("test_session_create_table.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("create", sessionId: sessionId, rows: 2, columns: 3);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Single(tables);
        Assert.Equal(2, tables[0].Rows.Count);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteTableInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_table.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Table to delete");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete", sessionId: sessionId, tableIndex: 0);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var tables = sessionDoc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Empty(tables);
    }

    [Fact]
    public void InsertRow_WithSessionId_ShouldInsertRowInMemory()
    {
        var docPath = CreateWordDocument("test_session_insert_row.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Existing Row");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("insert_row", sessionId: sessionId, tableIndex: 0, rowIndex: 0);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var tables = sessionDoc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Equal(2, tables[0].Rows.Count);
    }

    [Fact]
    public void DeleteRow_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_row.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Row 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Row 2");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete_row", sessionId: sessionId, tableIndex: 0, rowIndex: 1);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var tables = sessionDoc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Single(tables[0].Rows);
    }

    [Fact]
    public void InsertColumn_WithSessionId_ShouldInsertInMemory()
    {
        var docPath = CreateWordDocument("test_session_insert_col.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("insert_column", sessionId: sessionId, tableIndex: 0, columnIndex: 0);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var tables = sessionDoc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Equal(2, tables[0].FirstRow.Cells.Count);
    }

    [Fact]
    public void DeleteColumn_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_col.docx");
        var builder = new DocumentBuilder(new Document());
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Col 1");
        builder.InsertCell();
        builder.Write("Col 2");
        builder.EndRow();
        builder.EndTable();
        builder.Document.Save(docPath);

        var sessionId = OpenSession(docPath);
        _tool.Execute("delete_column", sessionId: sessionId, tableIndex: 0, columnIndex: 1);

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var tables = sessionDoc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        Assert.Single(tables[0].FirstRow.Cells);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_table_path.docx");
        var builder1 = new DocumentBuilder(new Document());
        builder1.StartTable();
        builder1.InsertCell();
        builder1.Write("PathTable");
        builder1.EndRow();
        builder1.EndTable();
        builder1.Document.Save(docPath1);

        var docPath2 = CreateWordDocument("test_table_session.docx");
        var builder2 = new DocumentBuilder(new Document());
        builder2.StartTable();
        builder2.InsertCell();
        builder2.Write("SessionTable");
        builder2.EndRow();
        builder2.EndTable();
        builder2.Document.Save(docPath2);

        var sessionId = OpenSession(docPath2);
        var result = _tool.Execute("get", docPath1, sessionId);

        var data = GetResultData<GetTablesWordResult>(result);
        Assert.Equal(1, data.Count);
    }

    #endregion
}
