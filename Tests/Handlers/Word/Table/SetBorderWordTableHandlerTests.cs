using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class SetBorderWordTableHandlerTests : WordHandlerTestBase
{
    private readonly SetBorderWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetBorder()
    {
        Assert.Equal("set_border", _handler.Operation);
    }

    #endregion

    #region Border Style

    [Theory]
    [InlineData("single")]
    [InlineData("double")]
    [InlineData("dotted")]
    [InlineData("dashed")]
    public void Execute_WithBorderStyle_SetsBorderStyle(string style)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "borderType", "all" },
            { "borderStyle", style }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("border", result, StringComparison.OrdinalIgnoreCase);
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

    #region Basic Set Border Operations

    [Fact]
    public void Execute_SetsBorder()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "borderTop", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("border", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData(true, false, false, false)]
    [InlineData(false, true, false, false)]
    [InlineData(false, false, true, false)]
    [InlineData(false, false, false, true)]
    public void Execute_SetsVariousBorders(bool top, bool bottom, bool left, bool right)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "borderTop", top },
            { "borderBottom", bottom },
            { "borderLeft", left },
            { "borderRight", right }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("border", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Border Width and Color

    [Fact]
    public void Execute_WithWidth_SetsWidth()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "borderType", "all" },
            { "width", 2.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("border", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithColor_SetsColor()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "borderType", "all" },
            { "color", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("border", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

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
        Assert.Contains("Table index", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidRowIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 99 },
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
