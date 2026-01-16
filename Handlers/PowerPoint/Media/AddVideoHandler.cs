using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Media;

/// <summary>
///     Handler for adding video to PowerPoint presentations.
/// </summary>
public class AddVideoHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add_video";

    /// <summary>
    ///     Adds a video file to a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: videoPath.
    ///     Optional: slideIndex (default: 0), x, y, width, height, playMode.
    /// </param>
    /// <returns>Success message with video details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractAddVideoParameters(parameters);

        SecurityHelper.ValidateFilePath(p.VideoPath, "videoPath", true);

        if (!File.Exists(p.VideoPath))
            throw new FileNotFoundException($"Video file not found: {p.VideoPath}", p.VideoPath);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        using var videoStream = new FileStream(p.VideoPath, FileMode.Open, FileAccess.Read);
        var video = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.ReadStreamAndRelease);
        var videoFrame = slide.Shapes.AddVideoFrame(p.X, p.Y, p.Width, p.Height, video);

        videoFrame.PlayMode = p.PlayMode.ToLower() switch
        {
            "onclick" => VideoPlayModePreset.OnClick,
            "auto" => VideoPlayModePreset.Auto,
            "inclick" => VideoPlayModePreset.InClickSequence,
            _ => VideoPlayModePreset.Auto
        };

        MarkModified(context);

        return Success(
            $"Video embedded into slide {p.SlideIndex} at position ({videoFrame.X:F0}, {videoFrame.Y:F0}) with dimensions {videoFrame.Width:F0}x{videoFrame.Height:F0}.");
    }

    /// <summary>
    ///     Extracts add video parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add video parameters.</returns>
    private static AddVideoParameters ExtractAddVideoParameters(OperationParameters parameters)
    {
        return new AddVideoParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<string>("videoPath"),
            parameters.GetOptional("x", 100f),
            parameters.GetOptional("y", 100f),
            parameters.GetOptional("width", 400f),
            parameters.GetOptional("height", 300f),
            parameters.GetOptional("playMode", "auto"));
    }

    /// <summary>
    ///     Record for holding add video parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="VideoPath">The video file path.</param>
    /// <param name="X">The X position.</param>
    /// <param name="Y">The Y position.</param>
    /// <param name="Width">The width.</param>
    /// <param name="Height">The height.</param>
    /// <param name="PlayMode">The play mode.</param>
    private record AddVideoParameters(
        int SlideIndex,
        string VideoPath,
        float X,
        float Y,
        float Width,
        float Height,
        string PlayMode);
}
