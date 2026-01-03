using Aspose.Slides;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from VideoFrame elements
/// </summary>
public class VideoFrameDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "Video";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is IVideoFrame;
    }

    /// <inheritdoc />
    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IVideoFrame video)
            return null;

        return new
        {
            contentType = video.EmbeddedVideo?.ContentType,
            playMode = video.PlayMode.ToString(),
            volume = video.Volume.ToString(),
            fullScreenMode = video.FullScreenMode,
            hideAtShowing = video.HideAtShowing,
            playLoopMode = video.PlayLoopMode,
            rewindVideo = video.RewindVideo,
            linkPathLong = video.LinkPathLong
        };
    }
}