using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class CopyWordTableHandlerTests : WordHandlerTestBase
{
    private readonly CopyWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_CopyTable()
    {
        Assert.Equal("copy_table", _handler.Operation);
    }

    #endregion

    #region Basic Copy Operations

    [Fact]
    public void Execute_CopiesTable()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("copied", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, GetTableCount(doc));
        AssertModified(context);
    }

    [Fact]
    public void Execute_CopiesTableStructure()
    {
        var doc = CreateDocumentWithTable(4, 5);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        var tables = GetTables(doc);
        Assert.Equal(tables[0].Rows.Count, tables[1].Rows.Count);
        Assert.Equal(tables[0].Rows[0].Cells.Count, tables[1].Rows[0].Cells.Count);
    }

    #endregion

    #region Table Index

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Execute_WithTableIndex_CopiesSpecificTable(int tableIndex)
    {
        var doc = CreateDocumentWithTables(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", tableIndex }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("copied", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(4, GetTableCount(doc));
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sourceTableIndex", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static int GetTableCount(Document doc)
    {
        return doc.GetChildNodes(NodeType.Table, true).Count;
    }

    private static List<Aspose.Words.Tables.Table> GetTables(Document doc)
    {
        return doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
    }

    private static Document CreateDocumentWithTable(int rows, int cols)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        for (var i = 0; i < rows; i++)
        {
            for (var j = 0; j < cols; j++)
            {
                builder.InsertCell();
                builder.Write($"R{i}C{j}");
            }

            builder.EndRow();
        }

        builder.EndTable();
        return doc;
    }

    private static Document CreateDocumentWithTables(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var t = 0; t < count; t++)
        {
            builder.StartTable();
            for (var i = 0; i < 2; i++)
            {
                for (var j = 0; j < 2; j++)
                {
                    builder.InsertCell();
                    builder.Write($"T{t}R{i}C{j}");
                }

                builder.EndRow();
            }

            builder.EndTable();
            builder.Writeln();
        }

        return doc;
    }

    #endregion
}
