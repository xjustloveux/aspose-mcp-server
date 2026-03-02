using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Media;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Media;

[SupportedOSPlatform("windows")]
public class AddVideoHandlerTests : PptHandlerTestBase
{
    private readonly AddVideoHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_AddVideo()
    {
        SkipIfNotWindows();
        Assert.Equal("add_video", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [SkippableFact]
    public void Execute_AddsVideo()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempVideoFile();
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "videoPath", tempFile }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.True(pres.Slides[0].Shapes.Count > initialShapeCount);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithSlideIndex_AddsVideoToSlide()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempVideoFile();
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "videoPath", tempFile },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.True(pres.Slides[0].Shapes.Count > initialShapeCount);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutVideoPath_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("videoPath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "videoPath", "/nonexistent/video.mp4" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempVideoFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "videoPath", tempFile },
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
