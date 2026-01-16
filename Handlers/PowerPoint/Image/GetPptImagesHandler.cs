using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for getting image information from PowerPoint slides.
/// </summary>
public class GetPptImagesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets image information from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex
    /// </param>
    /// <returns>JSON string containing image information including count, positions, and sizes.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var pictures = PptImageHelper.GetPictureFrames(slide);

        var imageInfoList = pictures.Select((pic, index) => new
        {
            imageIndex = index,
            x = pic.X,
            y = pic.Y,
            width = pic.Width,
            height = pic.Height,
            contentType = pic.PictureFormat.Picture.Image?.ContentType ?? "unknown"
        }).ToList();

        var result = new
        {
            slideIndex = p.SlideIndex,
            imageCount = pictures.Count,
            images = imageInfoList
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts get parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get parameters.</returns>
    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(parameters.GetRequired<int>("slideIndex"));
    }

    /// <summary>
    ///     Record for holding get images parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    private record GetParameters(int SlideIndex);
}
