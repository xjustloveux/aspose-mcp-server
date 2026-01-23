using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class SetColumnWidthWordTableHandlerTests : WordHandlerTestBase
{
    private readonly SetColumnWidthWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetColumnWidth()
    {
        Assert.Equal("set_column_width", _handler.Operation);
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

    #region Basic Set Width Operations

    [Fact]
    public void Execute_SetsColumnWidth()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 1 },
            { "columnWidth", 100.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("width", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_SetsWidthForVariousColumns(int columnIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", columnIndex },
            { "columnWidth", 150.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("width", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(50.0)]
    [InlineData(100.0)]
    [InlineData(200.0)]
    public void Execute_SetsVariousWidths(double width)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "columnWidth", width }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"{width}", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnWidth", 100.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutColumnWidth_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnWidth", ex.Message);
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
            { "columnIndex", invalidIndex },
            { "columnWidth", 100.0 }
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
            { "columnWidth", 100.0 },
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Table index", ex.Message);
    }

    #endregion
}
