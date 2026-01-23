using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class MergeCellsWordTableHandlerTests : WordHandlerTestBase
{
    private readonly MergeCellsWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_MergeCells()
    {
        Assert.Equal("merge_cells", _handler.Operation);
    }

    #endregion

    #region Helper Methods

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

    #endregion

    #region Basic Merge Operations

    [Fact]
    public void Execute_MergesCells()
    {
        var doc = CreateDocumentWithTable(4, 4);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startRow", 0 },
            { "endRow", 1 },
            { "startCol", 0 },
            { "endCol", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_MergesHorizontally()
    {
        var doc = CreateDocumentWithTable(3, 4);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startRow", 0 },
            { "endRow", 0 },
            { "startCol", 0 },
            { "endCol", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_MergesVertically()
    {
        var doc = CreateDocumentWithTable(4, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startRow", 0 },
            { "endRow", 2 },
            { "startCol", 0 },
            { "endCol", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("merged", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingParameters_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startRow", 0 },
            { "endRow", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("required", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startRow", 0 },
            { "endRow", 1 },
            { "startCol", 0 },
            { "endCol", 1 },
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Table index", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidRowRange_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startRow", 99 },
            { "endRow", 99 },
            { "startCol", 0 },
            { "endCol", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
