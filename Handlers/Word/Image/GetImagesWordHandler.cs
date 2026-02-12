using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Image;
using WordParagraph = Aspose.Words.Paragraph;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Represents the position and context information for an image.
/// </summary>
/// <param name="Alignment">The alignment of the image.</param>
/// <param name="Position">The position information for the image.</param>
/// <param name="ContextText">The context text near the image.</param>
internal record ImagePositionContext(string? Alignment, WordImagePosition? Position, string? ContextText);

/// <summary>
///     Represents image data information including type and original size.
/// </summary>
/// <param name="ImageType">The type of the image.</param>
/// <param name="OriginalSize">The original size of the image.</param>
internal record ImageDataInfo(string? ImageType, ImageSize? OriginalSize);

/// <summary>
///     Handler for getting all images from Word documents.
/// </summary>
[ResultType(typeof(GetImagesWordResult))]
public class GetImagesWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all images from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sectionIndex (-1 for all sections)
    /// </param>
    /// <returns>A JSON string containing image information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetImagesParameters(parameters);
        var doc = context.Document;
        var shapes = WordImageHelper.GetAllImages(doc, p.SectionIndex);

        if (shapes.Count == 0)
            return CreateEmptyResult(p.SectionIndex);

        var imageList = BuildImageList(shapes);

        var result = new GetImagesWordResult
        {
            Count = shapes.Count,
            SectionIndex = p.SectionIndex == -1 ? null : p.SectionIndex,
            Images = imageList
        };

        return result;
    }

    /// <summary>
    ///     Extracts get images parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get images parameters.</returns>
    private static GetImagesParameters ExtractGetImagesParameters(OperationParameters parameters)
    {
        return new GetImagesParameters(parameters.GetOptional("sectionIndex", 0));
    }

    /// <summary>
    ///     Creates an empty result when no images are found.
    /// </summary>
    /// <param name="sectionIndex">The section index searched.</param>
    /// <returns>An object with the empty result.</returns>
    private static GetImagesWordResult CreateEmptyResult(int sectionIndex)
    {
        var emptyResult = new GetImagesWordResult
        {
            Count = 0,
            SectionIndex = sectionIndex == -1 ? null : sectionIndex,
            Images = Array.Empty<WordImageInfo>(),
            Message = sectionIndex == -1
                ? "No images found in document"
                : $"No images found in section {sectionIndex}, use sectionIndex=-1 to search all sections"
        };
        return emptyResult;
    }

    /// <summary>
    ///     Builds the list of image information objects.
    /// </summary>
    /// <param name="shapes">The list of image shapes.</param>
    /// <returns>A list of image information objects.</returns>
    private static List<WordImageInfo> BuildImageList(List<WordShape> shapes)
    {
        List<WordImageInfo> imageList = [];
        for (var i = 0; i < shapes.Count; i++) imageList.Add(BuildImageInfo(shapes[i], i));

        return imageList;
    }

    /// <summary>
    ///     Builds the information record for a single image.
    /// </summary>
    /// <param name="shape">The image shape.</param>
    /// <param name="index">The image index.</param>
    /// <returns>A WordImageInfo record containing image information.</returns>
    private static WordImageInfo BuildImageInfo(WordShape shape, int index)
    {
        var positionContext = GetPositionAndContext(shape);
        var imageDataInfo = GetImageDataInfo(shape);

        return new WordImageInfo
        {
            Index = index,
            Name = string.IsNullOrEmpty(shape.Name) ? null : shape.Name,
            Width = shape.Width,
            Height = shape.Height,
            IsInline = shape.IsInline,
            Alignment = positionContext.Alignment,
            Position = positionContext.Position,
            Context = positionContext.ContextText,
            ImageType = imageDataInfo.ImageType,
            OriginalSize = imageDataInfo.OriginalSize,
            Hyperlink = string.IsNullOrEmpty(shape.HRef) ? null : shape.HRef,
            AltText = string.IsNullOrEmpty(shape.AlternativeText) ? null : shape.AlternativeText,
            Title = string.IsNullOrEmpty(shape.Title) ? null : shape.Title
        };
    }

    /// <summary>
    ///     Gets the position and context information for an image.
    /// </summary>
    /// <param name="shape">The image shape.</param>
    /// <returns>An ImagePositionContext containing alignment, position, and context text.</returns>
    private static ImagePositionContext GetPositionAndContext(
        WordShape shape)
    {
        if (shape.IsInline)
            return GetInlinePositionAndContext(shape);

        return GetFloatingPositionAndContext(shape);
    }

    /// <summary>
    ///     Gets position and context for an inline image.
    /// </summary>
    /// <param name="shape">The inline image shape.</param>
    /// <returns>An ImagePositionContext containing alignment, position, and context text.</returns>
    private static ImagePositionContext GetInlinePositionAndContext(
        WordShape shape)
    {
        if (shape.ParentNode is WordParagraph parentPara)
        {
            var alignment = parentPara.ParagraphFormat.Alignment.ToString();
            var contextText = TruncateText(parentPara.GetText().Trim());
            return new ImagePositionContext(alignment, null, contextText);
        }

        return new ImagePositionContext(null, new WordImagePosition { X = shape.Left, Y = shape.Top }, null);
    }

    /// <summary>
    ///     Gets position and context for a floating image.
    /// </summary>
    /// <param name="shape">The floating image shape.</param>
    /// <returns>An ImagePositionContext containing alignment, position, and context text.</returns>
    private static ImagePositionContext GetFloatingPositionAndContext(
        WordShape shape)
    {
        var position = new WordImagePosition
        {
            X = Math.Round(shape.Left, 1),
            Y = Math.Round(shape.Top, 1),
            HorizontalAlignment = shape.HorizontalAlignment.ToString(),
            VerticalAlignment = shape.VerticalAlignment.ToString(),
            WrapType = shape.WrapType.ToString()
        };

        string? contextText = null;
        if (shape.GetAncestor(NodeType.Paragraph) is WordParagraph nearestPara)
            contextText = TruncateText(nearestPara.GetText().Trim());

        return new ImagePositionContext(null, position, contextText);
    }

    /// <summary>
    ///     Truncates text to a maximum length.
    /// </summary>
    /// <param name="text">The text to truncate.</param>
    /// <returns>The truncated text or null if empty.</returns>
    private static string? TruncateText(string text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        return text.Length > 30 ? text[..30] + "..." : text;
    }

    /// <summary>
    ///     Gets image data information from the shape.
    /// </summary>
    /// <param name="shape">The image shape.</param>
    /// <returns>An ImageDataInfo containing image type and original size.</returns>
    private static ImageDataInfo GetImageDataInfo(WordShape shape)
    {
        if (shape.ImageData == null) return new ImageDataInfo(null, null);

        var imageType = shape.ImageData.ImageType.ToString();
        var imageSize = shape.ImageData.ImageSize;
        var originalSize = new ImageSize
        {
            WidthPixels = imageSize.WidthPixels,
            HeightPixels = imageSize.HeightPixels
        };

        return new ImageDataInfo(imageType, originalSize);
    }

    /// <summary>
    ///     Record to hold get images parameters.
    /// </summary>
    private sealed record GetImagesParameters(int SectionIndex);
}
