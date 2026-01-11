using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.DataOperations;

public class GetStatisticsHandlerTests : PptHandlerTestBase
{
    private readonly GetStatisticsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetStatistics()
    {
        Assert.Equal("get_statistics", _handler.Operation);
    }

    #endregion

    #region Basic Get Statistics Operations

    [Fact]
    public void Execute_ReturnsSlideCount()
    {
        var presentation = CreatePresentationWithSlides(3);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"totalSlides\": 3", result);
    }

    [Fact]
    public void Execute_ReturnsShapeCount()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("totalShapes", result);
    }

    [Fact]
    public void Execute_ReturnsSlideSizeInfo()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slideSize", result);
        Assert.Contains("width", result);
        Assert.Contains("height", result);
    }

    [Fact]
    public void Execute_ReturnsMediaCounts()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("totalImages", result);
        Assert.Contains("totalAudio", result);
        Assert.Contains("totalVideo", result);
    }

    [Fact]
    public void Execute_ReturnsLayoutAndMasterCounts()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("totalLayouts", result);
        Assert.Contains("totalMasters", result);
    }

    #endregion
}
