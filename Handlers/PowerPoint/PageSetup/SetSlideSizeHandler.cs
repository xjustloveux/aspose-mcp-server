using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.PageSetup;

/// <summary>
///     Handler for setting slide size in PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetSlideSizeHandler : OperationHandlerBase<Presentation>
{
    private const float MinSizePoints = 1f;
    private const float MaxSizePoints = 5000f;

    /// <inheritdoc />
    public override string Operation => "set_size";

    /// <summary>
    ///     Sets the slide size for the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: preset (OnScreen16x9, Widescreen, OnScreen16x10, Letter, A4, Banner, Custom),
    ///     width, height (required when preset=Custom),
    ///     scaleType (EnsureFit, Maximize, DoNotScale; default: EnsureFit).
    /// </param>
    /// <returns>Success message with slide size information.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the preset is unsupported, custom size is missing width/height, or scaleType is invalid.
    /// </exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSetSlideSizeParameters(parameters);
        var presentation = context.Document;
        var slideSize = presentation.SlideSize;

        var type = ParseSlideSizeType(p.Preset);
        var scaleType = ParseScaleType(p.ScaleType);

        if (type == SlideSizeType.Custom)
        {
            if (!p.Width.HasValue || !p.Height.HasValue)
                throw new ArgumentException("Custom size requires width and height.");

            ValidateSizeRange(p.Width.Value, p.Height.Value);
            slideSize.SetSize((float)p.Width.Value, (float)p.Height.Value, scaleType);
        }
        else
        {
            slideSize.SetSize(type, scaleType);
        }

        MarkModified(context);

        var sizeInfo = slideSize.Type == SlideSizeType.Custom
            ? $" ({slideSize.Size.Width}x{slideSize.Size.Height})"
            : "";

        return new SuccessResult { Message = $"Slide size set to {slideSize.Type}{sizeInfo}." };
    }

    /// <summary>
    ///     Validates that the width and height values are within acceptable range.
    /// </summary>
    /// <param name="width">The width value in points.</param>
    /// <param name="height">The height value in points.</param>
    /// <exception cref="ArgumentException">Thrown when width or height is outside the valid range.</exception>
    private static void ValidateSizeRange(double width, double height)
    {
        if (width < MinSizePoints || width > MaxSizePoints)
            throw new ArgumentException($"Width must be between {MinSizePoints} and {MaxSizePoints} points.");

        if (height < MinSizePoints || height > MaxSizePoints)
            throw new ArgumentException($"Height must be between {MinSizePoints} and {MaxSizePoints} points.");
    }

    /// <summary>
    ///     Parses a preset string to a <see cref="SlideSizeType" /> enum value.
    /// </summary>
    /// <param name="preset">The preset string (case-insensitive).</param>
    /// <returns>The parsed SlideSizeType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the preset string is not recognized.</exception>
    private static SlideSizeType ParseSlideSizeType(string preset)
    {
        return preset.ToLowerInvariant() switch
        {
            "onscreen16x9" => SlideSizeType.OnScreen16x9,
            "widescreen" => SlideSizeType.Widescreen,
            "onscreen16x10" => SlideSizeType.OnScreen16x10,
            "letter" => SlideSizeType.LetterPaper,
            "a4" => SlideSizeType.A4Paper,
            "banner" => SlideSizeType.Banner,
            "custom" => SlideSizeType.Custom,
            _ => throw new ArgumentException(
                $"Unsupported preset: {preset}. " +
                "Supported presets: OnScreen16x9, Widescreen, OnScreen16x10, Letter, A4, Banner, Custom.")
        };
    }

    /// <summary>
    ///     Parses a scale type string to a <see cref="SlideSizeScaleType" /> enum value.
    /// </summary>
    /// <param name="scaleType">The scale type string (case-insensitive).</param>
    /// <returns>The parsed SlideSizeScaleType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the scale type string is not recognized.</exception>
    private static SlideSizeScaleType ParseScaleType(string scaleType)
    {
        return scaleType.ToLowerInvariant() switch
        {
            "ensurefit" => SlideSizeScaleType.EnsureFit,
            "maximize" => SlideSizeScaleType.Maximize,
            "donotscale" => SlideSizeScaleType.DoNotScale,
            _ => throw new ArgumentException(
                $"Unsupported scaleType: {scaleType}. " +
                "Supported values: EnsureFit, Maximize, DoNotScale.")
        };
    }

    private static SetSlideSizeParameters ExtractSetSlideSizeParameters(OperationParameters parameters)
    {
        return new SetSlideSizeParameters(
            parameters.GetOptional("preset", "OnScreen16x9"),
            parameters.GetOptional<double?>("width"),
            parameters.GetOptional<double?>("height"),
            parameters.GetOptional("scaleType", "EnsureFit"));
    }

    /// <param name="Preset">The slide size preset.</param>
    /// <param name="Width">The custom width in points.</param>
    /// <param name="Height">The custom height in points.</param>
    /// <param name="ScaleType">The content scale type (EnsureFit, Maximize, DoNotScale).</param>
    private sealed record SetSlideSizeParameters(string Preset, double? Width, double? Height, string ScaleType);
}
