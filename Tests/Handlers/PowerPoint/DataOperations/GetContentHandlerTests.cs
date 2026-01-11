using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"totalSlides\": 2", result);
    }

    [Fact]
    public void Execute_ReturnsSlideContent()
    {
        var presentation = CreatePresentationWithText("Hello World");
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slides", result);
        Assert.Contains("textContent", result);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"index\":", result);
    }

    [Fact]
    public void Execute_ReturnsHiddenStatus()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("hidden", result);
    }

    #endregion
}
