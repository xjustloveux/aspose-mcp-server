using AsposeMcpServer.Handlers.Word.Styles;
using AsposeMcpServer.Results.Word.Styles;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Styles;

public class GetWordStylesHandlerTests : WordHandlerTestBase
{
    private readonly GetWordStylesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetStyles()
    {
        Assert.Equal("get_styles", _handler.Operation);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsStyles()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStylesResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.ParagraphStyles);
        Assert.True(result.Count >= 0);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithIncludeBuiltIn_ReturnsAllStyles()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeBuiltIn", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStylesResult>(res);

        Assert.NotNull(result);
        Assert.True(result.IncludeBuiltIn);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStylesResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.ParagraphStyles);
    }

    #endregion
}
