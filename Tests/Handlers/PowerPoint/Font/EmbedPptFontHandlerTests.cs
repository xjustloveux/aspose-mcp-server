using AsposeMcpServer.Handlers.PowerPoint.Font;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Font;

public class EmbedPptFontHandlerTests : PptHandlerTestBase
{
    private readonly EmbedPptFontHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Embed()
    {
        Assert.Equal("embed", _handler.Operation);
    }

    #endregion

    #region Embed Font

    [Fact]
    public void Execute_WithExistingFont_ReturnsSuccessResult()
    {
        var pres = CreatePresentationWithText("Hello World");
        var allFonts = pres.FontsManager.GetFonts();
        Assert.NotEmpty(allFonts);
        var fontName = allFonts[0].FontName;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", fontName }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains(fontName, result.Message);
    }

    [Fact]
    public void Execute_WithNonExistentFont_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", "NonExistentFontXYZ12345" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithMissingFontName_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithSubsetMode_ReturnsSuccessResult()
    {
        var pres = CreatePresentationWithText("Hello World");
        var allFonts = pres.FontsManager.GetFonts();
        Assert.NotEmpty(allFonts);
        var fontName = allFonts[0].FontName;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", fontName },
            { "embedMode", "subset" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains(fontName, result.Message);
    }

    #endregion
}
