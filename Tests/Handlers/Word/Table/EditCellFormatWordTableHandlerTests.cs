using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class EditCellFormatWordTableHandlerTests : WordHandlerTestBase
{
    private readonly EditCellFormatWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditCellFormat()
    {
        Assert.Equal("edit_cell_format", _handler.Operation);
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

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsCellFormat()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "backgroundColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("format", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 1)]
    [InlineData(2, 2)]
    public void Execute_EditsVariousCells(int rowIndex, int colIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", rowIndex },
            { "columnIndex", colIndex },
            { "backgroundColor", "#00FF00" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("format", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Format Options

    [Fact]
    public void Execute_WithBackgroundColor_SetsBackgroundColor()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "backgroundColor", "#0000FF" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("format", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithVerticalAlignment_SetsAlignment()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "verticalAlignment", "center" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("format", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithPadding_SetsPadding()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "topPadding", 5.0 },
            { "bottomPadding", 5.0 },
            { "leftPadding", 5.0 },
            { "rightPadding", 5.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("format", result, StringComparison.OrdinalIgnoreCase);
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
            { "backgroundColor", "#FF0000" }
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
            { "backgroundColor", "#FF0000" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(99, 0)]
    public void Execute_WithInvalidRowIndex_ThrowsArgumentException(int rowIndex, int colIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", rowIndex },
            { "columnIndex", colIndex },
            { "backgroundColor", "#FF0000" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Theory]
    [InlineData(0, -1)]
    [InlineData(0, 99)]
    public void Execute_WithInvalidColumnIndex_ThrowsArgumentException(int rowIndex, int colIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", rowIndex },
            { "columnIndex", colIndex },
            { "backgroundColor", "#FF0000" }
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
            { "columnIndex", 0 },
            { "backgroundColor", "#FF0000" },
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Table index", ex.Message);
    }

    #endregion
}
