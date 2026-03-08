using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Results.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.DataOperations;

[SupportedOSPlatform("windows")]
public class GetContentHandlerTests : PptHandlerTestBase
{
    private readonly GetContentHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetContent()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Basic Get Content Operations

    [SkippableFact]
    public void Execute_ReturnsTotalSlides()
    {
        SkipIfNotWindows();
        var presentation = CreatePresentationWithSlides(2);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentPptResult>(res);

        Assert.Equal(2, result.TotalSlides);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideContent()
    {
        SkipIfNotWindows();
        var presentation = CreatePresentationWithText("Hello World");
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentPptResult>(res);

        Assert.NotNull(result.Slides);
        Assert.True(result.Slides.Count > 0);
        Assert.NotNull(result.Slides[0].TextContent);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentPptResult>(res);

        Assert.NotNull(result.Slides);
        Assert.True(result.Slides.Count > 0);
        Assert.Equal(0, result.Slides[0].Index);
    }

    [SkippableFact]
    public void Execute_ReturnsHiddenStatus()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentPptResult>(res);

        Assert.NotNull(result.Slides);
        Assert.True(result.Slides.Count > 0);
        Assert.False(result.Slides[0].Hidden);
    }

    #endregion
}
