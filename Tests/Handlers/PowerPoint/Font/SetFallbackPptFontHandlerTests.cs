using AsposeMcpServer.Handlers.PowerPoint.Font;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Font;

public class SetFallbackPptFontHandlerTests : PptHandlerTestBase
{
    private readonly SetFallbackPptFontHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFallback()
    {
        Assert.Equal("set_fallback", _handler.Operation);
    }

    #endregion

    #region Set Fallback

    [Fact]
    public void Execute_WithFallbackFont_ReturnsSuccessResult()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fallbackFont", "Arial" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Arial", result.Message);
    }

    [Fact]
    public void Execute_WithFallbackFont_MarksContextModified()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fallbackFont", "Arial" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomUnicodeRange_ReturnsSuccessResult()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fallbackFont", "Arial" },
            { "unicodeStart", 0x4E00 },
            { "unicodeEnd", 0x9FFF }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Arial", result.Message);
        Assert.Contains("4E00", result.Message);
        Assert.Contains("9FFF", result.Message);
    }

    [Fact]
    public void Execute_WithMissingFallbackFont_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_AddsFallbackRule()
    {
        var pres = CreatePresentationWithText("Hello World");
        var initialCount = pres.FontsManager.FontFallBackRulesCollection.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fallbackFont", "Arial" }
        });

        _handler.Execute(context, parameters);

        Assert.True(pres.FontsManager.FontFallBackRulesCollection.Count > initialCount);
    }

    #endregion
}
