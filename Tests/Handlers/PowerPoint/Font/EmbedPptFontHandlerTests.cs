using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Font;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Font;

[SupportedOSPlatform("windows")]
public class EmbedPptFontHandlerTests : PptHandlerTestBase
{
    private readonly EmbedPptFontHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Embed()
    {
        SkipIfNotWindows();
        Assert.Equal("embed", _handler.Operation);
    }

    #endregion

    #region Embed Font

    [SkippableFact]
    public void Execute_WithExistingFont_ReturnsSuccessResult()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithNonExistentFont_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", "NonExistentFontXYZ12345" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithMissingFontName_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithSubsetMode_ReturnsSuccessResult()
    {
        SkipIfNotWindows();
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
