using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Media;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Media;

[SupportedOSPlatform("windows")]
public class AddAudioHandlerTests : PptHandlerTestBase
{
    private readonly AddAudioHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_AddAudio()
    {
        SkipIfNotWindows();
        Assert.Equal("add_audio", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [SkippableFact]
    public void Execute_AddsAudio()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempAudioFile();
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", tempFile }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.True(pres.Slides[0].Shapes.Count > initialShapeCount);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithSlideIndex_AddsAudioToSlide()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempAudioFile();
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", tempFile },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.True(pres.Slides[0].Shapes.Count > initialShapeCount);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithPosition_AddsAudioAtPosition()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempAudioFile();
        var pres = CreateEmptyPresentation();
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", tempFile },
            { "x", 200f },
            { "y", 150f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            Assert.True(pres.Slides[0].Shapes.Count > initialShapeCount);
            var lastShape = pres.Slides[0].Shapes[^1];
            Assert.Equal(200f, lastShape.X);
            Assert.Equal(150f, lastShape.Y);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutAudioPath_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("audioPath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", "/nonexistent/audio.mp3" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempAudioFile();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "audioPath", tempFile },
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
