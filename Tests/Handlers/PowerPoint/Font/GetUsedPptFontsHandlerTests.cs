using AsposeMcpServer.Handlers.PowerPoint.Font;
using AsposeMcpServer.Results.PowerPoint.Font;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Font;

public class GetUsedPptFontsHandlerTests : PptHandlerTestBase
{
    private readonly GetUsedPptFontsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetUsed()
    {
        Assert.Equal("get_used", _handler.Operation);
    }

    #endregion

    #region Get Used Fonts

    [Fact]
    public void Execute_ReturnsGetFontsPptResult()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<GetFontsPptResult>(res);
    }

    [Fact]
    public void Execute_ReturnsNonEmptyFontList()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.True(result.Count > 0);
        Assert.NotEmpty(result.Items);
    }

    [Fact]
    public void Execute_ReturnsFontWithName()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.All(result.Items, item => Assert.NotNull(item.FontName));
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.Equal(result.Items.Count, result.Count);
    }

    [Fact]
    public void Execute_ReturnsMessage()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.NotNull(result.Message);
        Assert.Contains("font", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_DoesNotModifyContext()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithEmptyPresentation_ReturnsResult()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.NotNull(result);
        Assert.Equal(result.Items.Count, result.Count);
    }

    [Fact]
    public void Execute_SystemFonts_MarkedAsNotCustom()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        var calibriFont = result.Items.FirstOrDefault(f =>
            f.FontName.Equals("Calibri", StringComparison.OrdinalIgnoreCase));
        if (calibriFont != null) Assert.False(calibriFont.IsCustom);
    }

    [Fact]
    public void Execute_ReturnsEmbeddedCount()
    {
        var pres = CreatePresentationWithText("Hello World");
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFontsPptResult>(res);
        Assert.True(result.EmbeddedCount >= 0);
    }

    #endregion
}
