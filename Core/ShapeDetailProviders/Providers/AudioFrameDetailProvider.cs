using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

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
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IAudioFrame audio)
            return null;

        return new AudioFrameDetails
        {
            ContentType = audio.EmbeddedAudio?.ContentType,
            PlayMode = audio.PlayMode.ToString(),
            Volume = audio.Volume.ToString(),
            PlayAcrossSlides = audio.PlayAcrossSlides,
            RewindAudio = audio.RewindAudio,
            HideAtShowing = audio.HideAtShowing,
            PlayLoopMode = audio.PlayLoopMode
        };
    }
}
