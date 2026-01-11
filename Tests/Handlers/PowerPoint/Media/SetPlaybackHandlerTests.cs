using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Media;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("onclick", result.ToLower());
        Assert.Equal(AudioPlayModePreset.OnClick, audioFrame.PlayMode);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("loop=true", result.ToLower());
        Assert.True(audioFrame.PlayLoopMode);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains(volumeStr, result.ToLower());
        Assert.Equal(expectedMode, audioFrame.Volume);
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

    #endregion
}
