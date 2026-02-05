using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.Font;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Font;

public class RemoveEmbeddedPptFontHandlerTests : PptHandlerTestBase
{
    private readonly RemoveEmbeddedPptFontHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_RemoveEmbedded()
    {
        Assert.Equal("remove_embedded", _handler.Operation);
    }

    #endregion

    #region Remove Embedded Font

    [Fact]
    public void Execute_WithNonEmbeddedFont_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", "Arial" }
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
    public void Execute_WithEmbeddedFont_RemovesIt()
    {
        var pres = CreatePresentationWithText("Hello World");
        var allFonts = pres.FontsManager.GetFonts();
        if (allFonts.Length == 0)
            return;

        var fontToEmbed = allFonts[0];
        pres.FontsManager.AddEmbeddedFont(fontToEmbed, EmbedFontCharacters.All);

        var embeddedBefore = pres.FontsManager.GetEmbeddedFonts().Length;
        Assert.True(embeddedBefore > 0);

        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontName", fontToEmbed.FontName }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains(fontToEmbed.FontName, result.Message);
        AssertModified(context);
    }

    #endregion
}
