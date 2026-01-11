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
        var slideIndex = parameters.GetRequired<int>("slideIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
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
            slideIndex,
            imageCount = pictures.Count,
            images = imageInfoList
        };

        return JsonResult(result);
    }
}
