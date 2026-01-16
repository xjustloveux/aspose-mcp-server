using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Background;

/// <summary>
///     Handler for getting PowerPoint slide background information.
/// </summary>
public class GetPptBackgroundHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets background information for a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndex (default: 0)
    /// </param>
    /// <returns>JSON string containing background information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetBackgroundParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var background = slide.Background;
        var fillFormat = background?.FillFormat;

        string? colorHex = null;
        double? opacity = null;

        if (fillFormat?.FillType == FillType.Solid)
            try
            {
                var solidColor = fillFormat.SolidFillColor.Color;
                if (!solidColor.IsEmpty)
                {
                    colorHex = solidColor.A < 255
                        ? $"#{solidColor.A:X2}{solidColor.R:X2}{solidColor.G:X2}{solidColor.B:X2}"
                        : $"#{solidColor.R:X2}{solidColor.G:X2}{solidColor.B:X2}";
                    opacity = Math.Round(solidColor.A / 255.0, 2);
                }
            }
            catch
            {
                // Ignore: color extraction may fail for non-solid fills or unsupported color formats
            }

        var result = new
        {
            slideIndex = p.SlideIndex,
            hasBackground = background != null,
            fillType = fillFormat?.FillType.ToString(),
            color = colorHex,
            opacity,
            isPictureFill = fillFormat?.FillType == FillType.Picture
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts get background parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get background parameters.</returns>
    private static GetBackgroundParameters ExtractGetBackgroundParameters(OperationParameters parameters)
    {
        return new GetBackgroundParameters(parameters.GetOptional("slideIndex", 0));
    }

    /// <summary>
    ///     Record for holding get background parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    private record GetBackgroundParameters(int SlideIndex);
}
