using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

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
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IVideoFrame video)
            return null;

        return new VideoFrameDetails
        {
            ContentType = video.EmbeddedVideo?.ContentType,
            PlayMode = video.PlayMode.ToString(),
            Volume = video.Volume.ToString(),
            FullScreenMode = video.FullScreenMode,
            HideAtShowing = video.HideAtShowing,
            PlayLoopMode = video.PlayLoopMode,
            RewindVideo = video.RewindVideo,
            LinkPathLong = video.LinkPathLong
        };
    }
}
