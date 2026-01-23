using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class InsertColumnWordTableHandlerTests : WordHandlerTestBase
{
    private readonly InsertColumnWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertColumn()
    {
        Assert.Equal("insert_column", _handler.Operation);
    }

    #endregion

    #region Basic Insert Operations

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_InsertsColumnAtVariousPositions(int columnIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", columnIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(4, GetFirstTable(doc).Rows[0].Cells.Count);
    }

    #endregion

    #region No Tables Scenario

    [Fact]
    public void Execute_WithNoTables_ThrowsArgumentException()
    {
        var doc = new Document();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Table index", ex.Message);
    }

    #endregion

    #region Insert At End

    [Fact]
    public void Execute_InsertAfterLastColumn_AppendsColumn()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 2 },
            { "insertBefore", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("index 3", result.Message);
        Assert.Equal(4, GetFirstTable(doc).Rows[0].Cells.Count);
    }

    #endregion

    #region Insert Before/After

    [Fact]
    public void Execute_WithInsertBeforeTrue_InsertsBeforeColumn()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 1 },
            { "insertBefore", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("index 1", result.Message);
    }

    [Fact]
    public void Execute_WithInsertBeforeFalse_InsertsAfterColumn()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 1 },
            { "insertBefore", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("index 2", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(99)]
    public void Execute_WithInvalidColumnIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", invalidIndex }
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
            { "columnIndex", 0 },
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

    #region Column Data Tests

    [Fact]
    public void Execute_WithColumnData_PopulatesCells()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var columnDataJson = "[\"Header\", \"Row1\", \"Row2\"]";
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "columnData", columnDataJson }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInvalidColumnDataJson_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "columnData", "invalid json" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Invalid columnData JSON", ex.Message);
    }

    [Fact]
    public void Execute_WithColumnDataWithLineBreaks_HandlesLineBreaks()
    {
        var doc = CreateDocumentWithTable(2, 2);
        var context = CreateContext(doc);
        var columnDataJson = "[\"Line1\\nLine2\", \"Single\"]";
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "columnData", columnDataJson }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Section Index Tests

    [Fact]
    public void Execute_WithValidSectionIndex_InsertsInSection()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "sectionIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "sectionIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Section index", ex.Message);
    }

    #endregion
}
