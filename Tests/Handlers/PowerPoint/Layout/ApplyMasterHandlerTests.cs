using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Layout;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

[SupportedOSPlatform("windows")]
public class ApplyMasterHandlerTests : PptHandlerTestBase
{
    private readonly ApplyMasterHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_ApplyMaster()
    {
        SkipIfNotWindows();
        Assert.Equal("apply_master", _handler.Operation);
    }

    #endregion

    #region Basic Apply Operations

    [SkippableFact]
    public void Execute_AppliesMaster()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var expectedLayout = pres.Masters[0].LayoutSlides[0];
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 },
            { "layoutIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(expectedLayout, pres.Slides[0].LayoutSlide);

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithSlideIndices_AppliesToSpecificSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var expectedLayout = pres.Masters[0].LayoutSlides[0];
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 },
            { "layoutIndex", 0 },
            { "slideIndices", "[0, 1]" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            Assert.Equal(expectedLayout, pres.Slides[0].LayoutSlide);
            Assert.Equal(expectedLayout, pres.Slides[1].LayoutSlide);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutMasterIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layoutIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithoutLayoutIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidMasterIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 99 },
            { "layoutIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidLayoutIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "masterIndex", 0 },
            { "layoutIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
