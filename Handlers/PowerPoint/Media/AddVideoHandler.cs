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
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var videoPath = parameters.GetRequired<string>("videoPath");
        var x = parameters.GetOptional("x", 100f);
        var y = parameters.GetOptional("y", 100f);
        var width = parameters.GetOptional("width", 400f);
        var height = parameters.GetOptional("height", 300f);
        var playMode = parameters.GetOptional("playMode", "auto");

        SecurityHelper.ValidateFilePath(videoPath, "videoPath", true);

        if (!File.Exists(videoPath))
            throw new FileNotFoundException($"Video file not found: {videoPath}", videoPath);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        using var videoStream = new FileStream(videoPath, FileMode.Open, FileAccess.Read);
        var video = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.ReadStreamAndRelease);
        var videoFrame = slide.Shapes.AddVideoFrame(x, y, width, height, video);

        videoFrame.PlayMode = playMode.ToLower() switch
        {
            "onclick" => VideoPlayModePreset.OnClick,
            "auto" => VideoPlayModePreset.Auto,
            "inclick" => VideoPlayModePreset.InClickSequence,
            _ => VideoPlayModePreset.Auto
        };

        MarkModified(context);

        return Success(
            $"Video embedded into slide {slideIndex} at position ({videoFrame.X:F0}, {videoFrame.Y:F0}) with dimensions {videoFrame.Width:F0}x{videoFrame.Height:F0}.");
    }
}
