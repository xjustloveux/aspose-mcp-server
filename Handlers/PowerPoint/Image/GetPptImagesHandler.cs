using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Image;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Handler for getting image information from PowerPoint slides.
/// </summary>
[ResultType(typeof(GetImagesPptResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var pictures = PptImageHelper.GetPictureFrames(slide);

        var imageInfoList = pictures.Select((pic, index) => new PptImageInfo
        {
            ImageIndex = index,
            X = pic.X,
            Y = pic.Y,
            Width = pic.Width,
            Height = pic.Height,
            ContentType = pic.PictureFormat.Picture.Image?.ContentType ?? "unknown"
        }).ToList();

        var result = new GetImagesPptResult
        {
            SlideIndex = p.SlideIndex,
            ImageCount = pictures.Count,
            Images = imageInfoList
        };

        return result;
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
    private sealed record GetParameters(int SlideIndex);
}
