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
        var p = ExtractAddAudioParameters(parameters);

        SecurityHelper.ValidateFilePath(p.AudioPath, "audioPath", true);

        if (!File.Exists(p.AudioPath))
            throw new FileNotFoundException($"Audio file not found: {p.AudioPath}", p.AudioPath);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        using var audioStream = new FileStream(p.AudioPath, FileMode.Open, FileAccess.Read);
        var audio = presentation.Audios.AddAudio(audioStream, LoadingStreamBehavior.ReadStreamAndRelease);
        var audioFrame = slide.Shapes.AddAudioFrameEmbedded(p.X, p.Y, p.Width, p.Height, audio);

        if (p.HideIcon)
            audioFrame.HideAtShowing = true;

        if (p.PlayAcrossSlides)
            audioFrame.PlayAcrossSlides = true;

        MarkModified(context);

        return Success(
            $"Audio embedded into slide {p.SlideIndex} at position ({audioFrame.X:F0}, {audioFrame.Y:F0}) with dimensions {audioFrame.Width:F0}x{audioFrame.Height:F0}.");
    }

    /// <summary>
    ///     Extracts add audio parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add audio parameters.</returns>
    private static AddAudioParameters ExtractAddAudioParameters(OperationParameters parameters)
    {
        return new AddAudioParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<string>("audioPath"),
            parameters.GetOptional("x", 100f),
            parameters.GetOptional("y", 100f),
            parameters.GetOptional("width", 50f),
            parameters.GetOptional("height", 50f),
            parameters.GetOptional("hideIcon", false),
            parameters.GetOptional("playAcrossSlides", false));
    }

    /// <summary>
    ///     Record for holding add audio parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="AudioPath">The audio file path.</param>
    /// <param name="X">The X position.</param>
    /// <param name="Y">The Y position.</param>
    /// <param name="Width">The width.</param>
    /// <param name="Height">The height.</param>
    /// <param name="HideIcon">Whether to hide the audio icon.</param>
    /// <param name="PlayAcrossSlides">Whether to play across slides.</param>
    private sealed record AddAudioParameters(
        int SlideIndex,
        string AudioPath,
        float X,
        float Y,
        float Width,
        float Height,
        bool HideIcon,
        bool PlayAcrossSlides);
}
