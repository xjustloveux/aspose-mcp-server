using AsposeMcpServer.Handlers.Word.Styles;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("paragraphStyles", result);
        Assert.Contains("count", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("includeBuiltIn", result);
        Assert.Contains("true", result);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
    }

    #endregion
}
