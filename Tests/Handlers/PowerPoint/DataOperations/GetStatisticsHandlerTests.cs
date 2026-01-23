using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Results.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.Equal(3, result.TotalSlides);
    }

    [Fact]
    public void Execute_ReturnsShapeCount()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.True(result.TotalShapes >= 0);
    }

    [Fact]
    public void Execute_ReturnsSlideSizeInfo()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.NotNull(result.SlideSize);
        Assert.True(result.SlideSize.Width > 0);
        Assert.True(result.SlideSize.Height > 0);
    }

    [Fact]
    public void Execute_ReturnsMediaCounts()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.True(result.TotalImages >= 0);
        Assert.True(result.TotalAudio >= 0);
        Assert.True(result.TotalVideo >= 0);
    }

    [Fact]
    public void Execute_ReturnsLayoutAndMasterCounts()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.True(result.TotalLayouts >= 0);
        Assert.True(result.TotalMasters >= 0);
    }

    #endregion
}
