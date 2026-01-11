using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class DeleteWordTableHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var doc = CreateDocumentWithTables(5);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Remaining", result);
        Assert.Contains("4", result);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesTable()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, GetTableCount(doc));
        AssertModified(context);
    }

    [Fact]
    public void Execute_DefaultIndex_DeletesFirstTable()
    {
        var doc = CreateDocumentWithTables(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(2, GetTableCount(doc));
    }

    #endregion

    #region Table Index

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithTableIndex_DeletesSpecificTable(int tableIndex)
    {
        var doc = CreateDocumentWithTables(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", tableIndex }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"#{tableIndex}", result);
        Assert.Equal(2, GetTableCount(doc));
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-10)]
    public void Execute_WithNegativeTableIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static int GetTableCount(Document doc)
    {
        return doc.GetChildNodes(NodeType.Table, true).Count;
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
