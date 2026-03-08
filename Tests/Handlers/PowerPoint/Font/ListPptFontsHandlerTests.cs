using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Font;
using AsposeMcpServer.Results.PowerPoint.Font;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Font;

[SupportedOSPlatform("windows")]
public class ListPptFontsHandlerTests : PptHandlerTestBase
{
    private readonly ListPptFontsHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetUsed()
    {
        SkipIfNotWindows();
        Assert.Equal("list", _handler.Operation);
    }

    #endregion

    #region Get Used Fonts

    [SkippableFact]
    public void Execute_ReturnsGetFontsPptResult()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<GetFontsPptResult>(res);
    }

    [SkippableFact]
    public void Execute_ReturnsNonEmptyFontList()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.True(result.Count > 0);
        Assert.NotEmpty(result.Items);
    }

    [SkippableFact]
    public void Execute_ReturnsFontWithName()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.All(result.Items, item => Assert.NotNull(item.FontName));
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.Equal(result.Items.Count, result.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsMessage()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.NotNull(result.Message);
        Assert.Contains("font", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_DoesNotModifyContext()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithEmptyPresentation_ReturnsResult()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.NotNull(result);
        Assert.Equal(result.Items.Count, result.Count);
    }

    [SkippableFact]
    public void Execute_SystemFonts_MarkedAsNotCustom()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        var calibriFont = result.Items.FirstOrDefault(f =>
            f.FontName.Equals("Calibri", StringComparison.OrdinalIgnoreCase));
        if (calibriFont != null) Assert.False(calibriFont.IsCustom);
    }

    [SkippableFact]
    public void Execute_ReturnsEmbeddedCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.True(result.EmbeddedCount >= 0);
    }

    #endregion
}
