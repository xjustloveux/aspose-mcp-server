using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Media;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Media;

public class SetPlaybackHandlerTests : PptHandlerTestBase
{
    private readonly SetPlaybackHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetPlayback()
    {
        Assert.Equal("set_playback", _handler.Operation);
    }

    #endregion

    #region Basic Playback Settings

    [Fact]
    public void Execute_SetsAudioPlaybackSettings()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var context = CreateContext(pres);
        var audioFrame = GetAudioFrame(pres);
        var shapeIndex = GetShapeIndex(pres.Slides[0], audioFrame);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex },
            { "playMode", "auto" },
            { "volume", "medium" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(AudioPlayModePreset.Auto, audioFrame.PlayMode);
        Assert.Equal(AudioVolumeMode.Medium, audioFrame.Volume);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithOnClickMode_SetsOnClickPlayback()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var context = CreateContext(pres);
        var audioFrame = GetAudioFrame(pres);
        var shapeIndex = GetShapeIndex(pres.Slides[0], audioFrame);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex },
            { "playMode", "onclick" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(AudioPlayModePreset.OnClick, audioFrame.PlayMode);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithLoop_EnablesLooping()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var context = CreateContext(pres);
        var audioFrame = GetAudioFrame(pres);
        var shapeIndex = GetShapeIndex(pres.Slides[0], audioFrame);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex },
            { "loop", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(audioFrame.PlayLoopMode);
        AssertModified(context);
    }

    [Theory]
    [InlineData("mute", AudioVolumeMode.Mute)]
    [InlineData("low", AudioVolumeMode.Low)]
    [InlineData("medium", AudioVolumeMode.Medium)]
    [InlineData("loud", AudioVolumeMode.Loud)]
    public void Execute_WithVolume_SetsVolumeLevel(string volumeStr, AudioVolumeMode expectedMode)
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var context = CreateContext(pres);
        var audioFrame = GetAudioFrame(pres);
        var shapeIndex = GetShapeIndex(pres.Slides[0], audioFrame);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex },
            { "volume", volumeStr }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(expectedMode, audioFrame.Volume);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidVolume_ThrowsArgumentException()
    {
        var tempFile = CreateTempAudioFile();
        var pres = CreatePresentationWithAudio(tempFile);
        var context = CreateContext(pres);
        var audioFrame = GetAudioFrame(pres);
        var shapeIndex = GetShapeIndex(pres.Slides[0], audioFrame);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex },
            { "volume", "invalid" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonMediaShape_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Test");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Video FullScreenMode

    [Fact]
    public void Execute_WithFullScreenModeTrue_SetsFullScreenOnVideo()
    {
        var tempFile = CreateTempVideoFile();
        var pres = CreatePresentationWithVideo(tempFile);
        var context = CreateContext(pres);
        var videoFrame = GetVideoFrame(pres);
        var shapeIndex = GetShapeIndex(pres.Slides[0], videoFrame);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex },
            { "fullScreenMode", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(videoFrame.FullScreenMode);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFullScreenModeFalse_SetsFullScreenOffOnVideo()
    {
        var tempFile = CreateTempVideoFile();
        var pres = CreatePresentationWithVideo(tempFile);
        var context = CreateContext(pres);
        var videoFrame = GetVideoFrame(pres);
        videoFrame.FullScreenMode = true;
        var shapeIndex = GetShapeIndex(pres.Slides[0], videoFrame);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex },
            { "fullScreenMode", false }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.False(videoFrame.FullScreenMode);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithoutFullScreenMode_DoesNotChangeFullScreen()
    {
        var tempFile = CreateTempVideoFile();
        var pres = CreatePresentationWithVideo(tempFile);
        var context = CreateContext(pres);
        var videoFrame = GetVideoFrame(pres);
        var originalValue = videoFrame.FullScreenMode;
        var shapeIndex = GetShapeIndex(pres.Slides[0], videoFrame);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", shapeIndex },
            { "playMode", "auto" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(originalValue, videoFrame.FullScreenMode);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithAudio(string audioPath)
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        using var audioStream = new FileStream(audioPath, FileMode.Open, FileAccess.Read);
        slide.Shapes.AddAudioFrameEmbedded(50, 50, 100, 100, audioStream);
        return pres;
    }

    private static IAudioFrame GetAudioFrame(Presentation pres)
    {
        return pres.Slides[0].Shapes.OfType<IAudioFrame>().First();
    }

    private static int GetShapeIndex(ISlide slide, IShape shape)
    {
        for (var i = 0; i < slide.Shapes.Count; i++)
            if (slide.Shapes[i] == shape)
                return i;
        return -1;
    }

    private static Presentation CreatePresentationWithVideo(string videoPath)
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        using var videoStream = new FileStream(videoPath, FileMode.Open, FileAccess.Read);
        var video = pres.Videos.AddVideo(videoStream, LoadingStreamBehavior.ReadStreamAndRelease);
        slide.Shapes.AddVideoFrame(50, 50, 320, 240, video);
        return pres;
    }

    private static IVideoFrame GetVideoFrame(Presentation pres)
    {
        return pres.Slides[0].Shapes.OfType<IVideoFrame>().First();
    }

    #endregion
}
