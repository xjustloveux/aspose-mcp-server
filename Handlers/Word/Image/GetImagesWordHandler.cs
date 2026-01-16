using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for getting all images from Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetImagesParameters(parameters);
        var doc = context.Document;
        var shapes = WordImageHelper.GetAllImages(doc, p.SectionIndex);

        if (shapes.Count == 0)
            return CreateEmptyResult(p.SectionIndex);

        var imageList = BuildImageList(shapes);

        var result = new
        {
            count = shapes.Count,
            sectionIndex = p.SectionIndex == -1 ? (int?)null : p.SectionIndex,
            images = imageList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
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
    /// <returns>A JSON string with the empty result.</returns>
    private static string CreateEmptyResult(int sectionIndex)
    {
        var emptyResult = new
        {
            count = 0,
            sectionIndex = sectionIndex == -1 ? (int?)null : sectionIndex,
            images = Array.Empty<object>(),
            message = sectionIndex == -1
                ? "No images found in document"
                : $"No images found in section {sectionIndex}, use sectionIndex=-1 to search all sections"
        };
        return JsonSerializer.Serialize(emptyResult, JsonDefaults.Indented);
    }

    /// <summary>
    ///     Builds the list of image information objects.
    /// </summary>
    /// <param name="shapes">The list of image shapes.</param>
    /// <returns>A list of image information objects.</returns>
    private static List<object> BuildImageList(List<WordShape> shapes)
    {
        List<object> imageList = [];
        for (var i = 0; i < shapes.Count; i++) imageList.Add(BuildImageInfo(shapes[i], i));

        return imageList;
    }

    /// <summary>
    ///     Builds the information dictionary for a single image.
    /// </summary>
    /// <param name="shape">The image shape.</param>
    /// <param name="index">The image index.</param>
    /// <returns>A dictionary containing image information.</returns>
    private static Dictionary<string, object?> BuildImageInfo(WordShape shape, int index)
    {
        var (alignment, position, contextText) = GetPositionAndContext(shape);

        var imageInfo = new Dictionary<string, object?>
        {
            ["index"] = index,
            ["name"] = string.IsNullOrEmpty(shape.Name) ? null : shape.Name,
            ["width"] = shape.Width,
            ["height"] = shape.Height,
            ["isInline"] = shape.IsInline
        };

        if (alignment != null) imageInfo["alignment"] = alignment;
        if (position != null) imageInfo["position"] = position;
        if (contextText != null) imageInfo["context"] = contextText;

        AddImageDataInfo(imageInfo, shape);
        AddOptionalProperties(imageInfo, shape);

        return imageInfo;
    }

    /// <summary>
    ///     Gets the position and context information for an image.
    /// </summary>
    /// <param name="shape">The image shape.</param>
    /// <returns>A tuple containing alignment, position, and context text.</returns>
    private static (string? alignment, object? position, string? contextText) GetPositionAndContext(WordShape shape)
    {
        if (shape.IsInline)
            return GetInlinePositionAndContext(shape);

        return GetFloatingPositionAndContext(shape);
    }

    /// <summary>
    ///     Gets position and context for an inline image.
    /// </summary>
    /// <param name="shape">The inline image shape.</param>
    /// <returns>A tuple containing alignment, position, and context text.</returns>
    private static (string? alignment, object? position, string? contextText) GetInlinePositionAndContext(
        WordShape shape)
    {
        if (shape.ParentNode is WordParagraph parentPara)
        {
            var alignment = parentPara.ParagraphFormat.Alignment.ToString();
            var contextText = TruncateText(parentPara.GetText().Trim());
            return (alignment, null, contextText);
        }

        return (null, new { x = shape.Left, y = shape.Top }, null);
    }

    /// <summary>
    ///     Gets position and context for a floating image.
    /// </summary>
    /// <param name="shape">The floating image shape.</param>
    /// <returns>A tuple containing alignment, position, and context text.</returns>
    private static (string? alignment, object? position, string? contextText) GetFloatingPositionAndContext(
        WordShape shape)
    {
        var position = new
        {
            x = Math.Round(shape.Left, 1),
            y = Math.Round(shape.Top, 1),
            horizontalAlignment = shape.HorizontalAlignment.ToString(),
            verticalAlignment = shape.VerticalAlignment.ToString(),
            wrapType = shape.WrapType.ToString()
        };

        string? contextText = null;
        if (shape.GetAncestor(NodeType.Paragraph) is WordParagraph nearestPara)
            contextText = TruncateText(nearestPara.GetText().Trim());

        return (null, position, contextText);
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
    ///     Adds image data information to the info dictionary.
    /// </summary>
    /// <param name="imageInfo">The image info dictionary.</param>
    /// <param name="shape">The image shape.</param>
    private static void AddImageDataInfo(Dictionary<string, object?> imageInfo, WordShape shape)
    {
        if (shape.ImageData == null) return;

        imageInfo["imageType"] = shape.ImageData.ImageType.ToString();
        var imageSize = shape.ImageData.ImageSize;
        imageInfo["originalSize"] = new { widthPixels = imageSize.WidthPixels, heightPixels = imageSize.HeightPixels };
    }

    /// <summary>
    ///     Adds optional properties to the info dictionary.
    /// </summary>
    /// <param name="imageInfo">The image info dictionary.</param>
    /// <param name="shape">The image shape.</param>
    private static void AddOptionalProperties(Dictionary<string, object?> imageInfo, WordShape shape)
    {
        if (!string.IsNullOrEmpty(shape.HRef)) imageInfo["hyperlink"] = shape.HRef;
        if (!string.IsNullOrEmpty(shape.AlternativeText)) imageInfo["altText"] = shape.AlternativeText;
        if (!string.IsNullOrEmpty(shape.Title)) imageInfo["title"] = shape.Title;
    }

    /// <summary>
    ///     Record to hold get images parameters.
    /// </summary>
    private sealed record GetImagesParameters(int SectionIndex);
}
