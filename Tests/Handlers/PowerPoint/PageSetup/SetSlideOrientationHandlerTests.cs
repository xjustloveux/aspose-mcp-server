using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.PageSetup;

[SupportedOSPlatform("windows")]
public class SetSlideOrientationHandlerTests : PptHandlerTestBase
{
    private readonly SetSlideOrientationHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_SetOrientation()
    {
        SkipIfNotWindows();
        Assert.Equal("set_orientation", _handler.Operation);
    }

    #endregion

    #region Basic Set Slide Orientation Operations

    [SkippableFact]
    public void Execute_SetsPortraitOrientation()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Portrait" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        var size = presentation.SlideSize.Size;
        Assert.True(size.Height > size.Width,
            $"Portrait orientation should have height ({size.Height}) > width ({size.Width})");
    }

    [SkippableFact]
    public void Execute_SetsLandscapeOrientation()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Landscape" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var size = presentation.SlideSize.Size;
        Assert.True(size.Width > size.Height,
            $"Landscape orientation should have width ({size.Width}) > height ({size.Height})");
    }

    [SkippableFact]
    public void Execute_ReturnsSizeInfo()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Portrait" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var size = presentation.SlideSize.Size;
        Assert.True(size.Width > 0, "Slide width should be positive");
        Assert.True(size.Height > 0, "Slide height should be positive");
        Assert.True(size.Height > size.Width,
            $"Portrait orientation should have height ({size.Height}) > width ({size.Width})");
    }

    #endregion
}
