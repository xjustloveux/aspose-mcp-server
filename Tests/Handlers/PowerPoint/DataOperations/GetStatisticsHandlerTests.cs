using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Results.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.DataOperations;

[SupportedOSPlatform("windows")]
public class GetStatisticsHandlerTests : PptHandlerTestBase
{
    private readonly GetStatisticsHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetStatistics()
    {
        SkipIfNotWindows();
        Assert.Equal("statistics", _handler.Operation);
    }

    #endregion

    #region Basic Get Statistics Operations

    [SkippableFact]
    public void Execute_ReturnsSlideCount()
    {
        SkipIfNotWindows();
        var presentation = CreatePresentationWithSlides(3);
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.Equal(3, result.TotalSlides);
    }

    [SkippableFact]
    public void Execute_ReturnsShapeCount()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.True(result.TotalShapes >= 0);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideSizeInfo()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.NotNull(result.SlideSize);
        Assert.True(result.SlideSize.Width > 0);
        Assert.True(result.SlideSize.Height > 0);
    }

    [SkippableFact]
    public void Execute_ReturnsMediaCounts()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStatisticsResult>(res);

        Assert.True(result.TotalImages >= 0);
        Assert.True(result.TotalAudio >= 0);
        Assert.True(result.TotalVideo >= 0);
    }

    [SkippableFact]
    public void Execute_ReturnsLayoutAndMasterCounts()
    {
        SkipIfNotWindows();
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
