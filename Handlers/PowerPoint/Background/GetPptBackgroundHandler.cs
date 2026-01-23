using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Background;

namespace AsposeMcpServer.Handlers.PowerPoint.Background;

/// <summary>
///     Handler for getting PowerPoint slide background information.
/// </summary>
[ResultType(typeof(GetBackgroundResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
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

        var result = new GetBackgroundResult
        {
            SlideIndex = p.SlideIndex,
            HasBackground = background != null,
            FillType = fillFormat?.FillType.ToString(),
            Color = colorHex,
            Opacity = opacity,
            IsPictureFill = fillFormat?.FillType == FillType.Picture
        };

        return result;
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
    private sealed record GetBackgroundParameters(int SlideIndex);
}
