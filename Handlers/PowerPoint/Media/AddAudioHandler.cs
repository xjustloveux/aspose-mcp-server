using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Media;

/// <summary>
///     Handler for adding audio to PowerPoint presentations.
/// </summary>
public class AddAudioHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add_audio";

    /// <summary>
    ///     Adds an audio file to a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: audioPath.
    ///     Optional: slideIndex (default: 0), x, y, width, height, hideIcon, playAcrossSlides.
    /// </param>
    /// <returns>Success message with audio details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var audioPath = parameters.GetRequired<string>("audioPath");
        var x = parameters.GetOptional("x", 100f);
        var y = parameters.GetOptional("y", 100f);
        var width = parameters.GetOptional("width", 50f);
        var height = parameters.GetOptional("height", 50f);
        var hideIcon = parameters.GetOptional("hideIcon", false);
        var playAcrossSlides = parameters.GetOptional("playAcrossSlides", false);

        SecurityHelper.ValidateFilePath(audioPath, "audioPath", true);

        if (!File.Exists(audioPath))
            throw new FileNotFoundException($"Audio file not found: {audioPath}", audioPath);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        using var audioStream = new FileStream(audioPath, FileMode.Open, FileAccess.Read);
        var audio = presentation.Audios.AddAudio(audioStream, LoadingStreamBehavior.ReadStreamAndRelease);
        var audioFrame = slide.Shapes.AddAudioFrameEmbedded(x, y, width, height, audio);

        if (hideIcon)
            audioFrame.HideAtShowing = true;

        if (playAcrossSlides)
            audioFrame.PlayAcrossSlides = true;

        MarkModified(context);

        return Success(
            $"Audio embedded into slide {slideIndex} at position ({audioFrame.X:F0}, {audioFrame.Y:F0}) with dimensions {audioFrame.Width:F0}x{audioFrame.Height:F0}.");
    }
}
