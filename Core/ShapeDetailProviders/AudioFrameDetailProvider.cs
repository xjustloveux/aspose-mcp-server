using Aspose.Slides;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from AudioFrame elements
/// </summary>
public class AudioFrameDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "Audio";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is IAudioFrame;
    }

    /// <inheritdoc />
    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IAudioFrame audio)
            return null;

        return new
        {
            contentType = audio.EmbeddedAudio?.ContentType,
            playMode = audio.PlayMode.ToString(),
            volume = audio.Volume.ToString(),
            playAcrossSlides = audio.PlayAcrossSlides,
            rewindAudio = audio.RewindAudio,
            hideAtShowing = audio.HideAtShowing,
            playLoopMode = audio.PlayLoopMode
        };
    }
}