using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class SplitCellWordTableHandlerTests : WordHandlerTestBase
{
    private readonly SplitCellWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SplitCell()
    {
        Assert.Equal("split_cell", _handler.Operation);
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

    #region Basic Split Operations

    [Fact]
    public void Execute_SplitsCell()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitRows", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var table = doc.Sections[0].Body.Tables[0];
            Assert.True(table.Rows.Count > 3);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_SplitsHorizontally()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitRows", 1 },
            { "splitCols", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var table = doc.Sections[0].Body.Tables[0];
            Assert.True(table.Rows[0].Cells.Count > 3);
        }
    }

    [Fact]
    public void Execute_SplitsVertically()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitRows", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var table = doc.Sections[0].Body.Tables[0];
            Assert.True(table.Rows.Count > 3);
        }
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRowIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "splitRows", 2 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rowIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "splitRows", 2 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitRows", 2 },
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Table index", ex.Message);
    }

    #endregion

    #region Split Budget

    /// <summary>
    ///     Splitting had no bound of its own: <c>splitCols</c> and <c>splitRows</c> were whatever
    ///     the caller sent, and each pair becomes a cell object (R3-R04).
    /// </summary>
    /// <param name="splitRows">Rows to split into.</param>
    /// <param name="splitCols">Columns to split into.</param>
    [Theory]
    [InlineData(100_000, 2)]
    [InlineData(2, 100_000)]
    [InlineData(50_000, 60)]
    public void Execute_WithASplitAboveTheBudget_ShouldBeRefused(int splitRows, int splitCols)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitRows", splitRows },
            { "splitCols", splitCols }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        var table = doc.GetChildNodes(NodeType.Table, true)
            .Cast<Aspose.Words.Tables.Table>().First();
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(3, table.Rows[0].Cells.Count);
    }

    /// <summary>
    ///     R4-R04: the budget was measured from the selected row alone, so a table whose other rows
    ///     were already past the column limit was priced as if they were as narrow as this one. A
    ///     Word table may be ragged; the widest row is what the format has to hold.
    /// </summary>
    [Fact]
    public void Execute_WhenAnotherRowIsAlreadyOverTheColumnLimit_ShouldBeRefused()
    {
        var doc = CreateDocumentWithRaggedTable(2, 70);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitCols", 2 }
        });

        var refusal = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columns", refusal.Message, StringComparison.OrdinalIgnoreCase);

        // Nothing was changed on the way to the refusal.
        var table = doc.GetChildNodes(NodeType.Table, true)
            .Cast<Aspose.Words.Tables.Table>().First();
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal(70, table.Rows[1].Cells.Count);
        AssertNotModified(context);
    }

    /// <summary>
    ///     Builds a table whose first row is narrow and whose second is as wide as asked. Aspose
    ///     holds more cells in a row than the Word format permits, which is how a document can
    ///     arrive already past the limit.
    /// </summary>
    /// <param name="narrowCells">Cells in the first row.</param>
    /// <param name="wideCells">Cells in the second row.</param>
    /// <returns>The document.</returns>
    private static Document CreateDocumentWithRaggedTable(int narrowCells, int wideCells)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        foreach (var cells in new[] { narrowCells, wideCells })
        {
            for (var c = 0; c < cells; c++)
            {
                builder.InsertCell();
                builder.Write($"c{c}");
            }

            builder.EndRow();
        }

        builder.EndTable();
        return doc;
    }

    [Fact]
    public void Execute_WithAnOrdinarySplit_ShouldStillWork()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitRows", 2 },
            { "splitCols", 2 }
        });

        _handler.Execute(context, parameters);

        var table = doc.GetChildNodes(NodeType.Table, true)
            .Cast<Aspose.Words.Tables.Table>().First();
        Assert.Equal(4, table.Rows[0].Cells.Count);
    }

    #endregion

    #region Requested Dimensions

    /// <summary>
    ///     Only the computed row width was checked, and the subtraction cancelled a zero out: a
    ///     split into zero columns added a row and reported success (R4-R04).
    /// </summary>
    /// <param name="splitRows">Rows to split into.</param>
    /// <param name="splitCols">Columns to split into.</param>
    [Theory]
    [InlineData(2, 0)]
    [InlineData(0, 2)]
    [InlineData(2, -1)]
    [InlineData(-1, 2)]
    public void Execute_WithADimensionBelowOne_ShouldBeRefused(int splitRows, int splitCols)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitRows", splitRows },
            { "splitCols", splitCols }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        var table = doc.GetChildNodes(NodeType.Table, true)
            .Cast<Aspose.Words.Tables.Table>().First();
        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(3, table.Rows[0].Cells.Count);
        AssertNotModified(context);
    }

    /// <summary>
    ///     The whole-table cap has to count the rows the table already has, not only the ones the
    ///     split adds (R4-R04).
    /// </summary>
    [Fact]
    public void Execute_WhenTheResultingTableWouldBeTooLarge_ShouldBeRefused()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "splitRows", 10_000 },
            { "splitCols", 2 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
