using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class DeleteRowWordTableHandlerTests : WordHandlerTestBase
{
    private readonly DeleteRowWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteRow()
    {
        Assert.Equal("delete_row", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesRowAtVariousPositions(int rowIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", rowIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, GetFirstTable(doc).Rows.Count);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsRemainingCount()
    {
        var doc = CreateDocumentWithTable(5, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Remaining", result.Message);
    }

    #endregion

    #region Cross-Section Flat Index (get→mutate consistency)

    [Fact]
    public void Execute_WithoutSectionIndex_UsesDocumentFlatTableIndexLikeGet()
    {
        var doc = CreateDocumentWithTablesInTwoSections();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 1 },
            { "rowIndex", 0 }
        });

        _handler.Execute(context, parameters);

        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        Assert.Equal(2, tables[0].Rows.Count);
        Assert.Equal(1, tables[1].Rows.Count);
    }

    private static Document CreateDocumentWithTablesInTwoSections()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("s0r0");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("s0r1");
        builder.EndRow();
        builder.EndTable();
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("s1r0");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("s1r1");
        builder.EndRow();
        builder.EndTable();
        return doc;
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRowIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rowIndex", ex.Message);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(99)]
    public void Execute_WithInvalidRowIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Table index", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Aspose.Words.Tables.Table GetFirstTable(Document doc)
    {
        return (Aspose.Words.Tables.Table)doc.GetChildNodes(NodeType.Table, true)[0];
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

    #endregion
}
