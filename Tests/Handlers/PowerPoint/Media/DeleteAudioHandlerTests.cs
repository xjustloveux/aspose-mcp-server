using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Media;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Media;

[SupportedOSPlatform("windows")]
public class DeleteAudioHandlerTests : PptHandlerTestBase
{
    private readonly DeleteAudioHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_DeleteAudio()
    {
        SkipIfNotWindows();
        Assert.Equal("delete_audio", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithAudio(string audioPath)
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        using var audioStream = new FileStream(audioPath, FileMode.Open, FileAccess.Read);
        var audio = pres.Audios.AddAudio(audioStream, LoadingStreamBehavior.ReadStreamAndRelease);
        slide.Shapes.AddAudioFrameEmbedded(100, 100, 50, 50, audio);
        return pres;
    }

    #endregion

    #region Basic Delete Operations

    [SkippableFact]
    public void Execute_DeletesAudio()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount - 1, pres.Slides[0].Shapes.Count);
        Assert.DoesNotContain(pres.Slides[0].Shapes, s => s is IAudioFrame);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_DeletesAudioFromSpecificSlide()
    {
        SkipIfNotWindows();
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var initialShapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialShapeCount - 1, pres.Slides[0].Shapes.Count);
        Assert.DoesNotContain(pres.Slides[0].Shapes, s => s is IAudioFrame);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableTheory]
    [InlineData(-1)]
    [InlineData(99)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", invalidIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithNonAudioShape_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Test");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not an audio frame", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
