using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class InsertRowWordTableHandlerTests : WordHandlerTestBase
{
    private readonly InsertRowWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertRow()
    {
        Assert.Equal("insert_row", _handler.Operation);
    }

    #endregion

    #region Basic Insert Operations

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_InsertsRowAtVariousPositions(int rowIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", rowIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(4, GetFirstTable(doc).Rows.Count);
    }

    #endregion

    #region Insert Before/After

    [Fact]
    public void Execute_WithInsertBeforeTrue_InsertsBeforeRow()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 1 },
            { "insertBefore", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("index 1", result);
        Assert.Equal(4, GetFirstTable(doc).Rows.Count);
    }

    [Fact]
    public void Execute_WithInsertBeforeFalse_InsertsAfterRow()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 1 },
            { "insertBefore", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("index 2", result);
        Assert.Equal(4, GetFirstTable(doc).Rows.Count);
    }

    #endregion

    #region Row Data

    [Fact]
    public void Execute_WithRowData_InsertsRowWithData()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "insertBefore", true },
            { "rowData", "[\"New1\", \"New2\", \"New3\"]" }
        });

        _handler.Execute(context, parameters);

        var table = GetFirstTable(doc);
        Assert.Equal(4, table.Rows.Count);
        var insertedRow = table.Rows[0];
        Assert.Equal("New1", GetCellText(insertedRow.Cells[0]));
        Assert.Equal("New2", GetCellText(insertedRow.Cells[1]));
        Assert.Equal("New3", GetCellText(insertedRow.Cells[2]));
    }

    [Fact]
    public void Execute_WithInvalidRowData_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "rowData", "invalid json" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Invalid rowData", ex.Message);
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

    private static string GetCellText(Cell cell)
    {
        return cell.GetText().Trim('\a', ' ', '\r', '\n');
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
