using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Results.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.DataOperations;

public class GetContentHandlerTests : PptHandlerTestBase
{
    private readonly GetContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetContent()
    {
        Assert.Equal("get_content", _handler.Operation);
    }

    #endregion

    #region Basic Get Content Operations

    [Fact]
    public void Execute_ReturnsTotalSlides()
    {
        var presentation = CreatePresentationWithSlides(2);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentPptResult>(res);

        Assert.Equal(2, result.TotalSlides);
    }

    [Fact]
    public void Execute_ReturnsSlideContent()
    {
        var presentation = CreatePresentationWithText("Hello World");
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentPptResult>(res);

        Assert.NotNull(result.Slides);
        Assert.True(result.Slides.Count > 0);
        Assert.NotNull(result.Slides[0].TextContent);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentPptResult>(res);

        Assert.NotNull(result.Slides);
        Assert.True(result.Slides.Count > 0);
        Assert.Equal(0, result.Slides[0].Index);
    }

    [Fact]
    public void Execute_ReturnsHiddenStatus()
    {
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
