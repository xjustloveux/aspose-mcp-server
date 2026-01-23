using AsposeMcpServer.Handlers.Word.Format;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Format;

public class SetParagraphBorderWordHandlerTests : WordHandlerTestBase
{
    private readonly SetParagraphBorderWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetParagraphBorder()
    {
        Assert.Equal("set_paragraph_border", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidBorderPosition_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "borderPosition", "invalid" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsParagraphBorder()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "borderTop", true },
            { "borderBottom", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("borders set", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("top", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("bottom", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithBorderPositionAll_SetsAllBorders()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "borderPosition", "all" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("top", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("bottom", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("left", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("right", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithBorderPositionNone_ClearsBorders()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "borderPosition", "none" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("none", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithCustomLineStyle_SetsBorderWithStyle()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "borderTop", true },
            { "lineStyle", "double" },
            { "lineWidth", 1.0 },
            { "lineColor", "FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("top", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
