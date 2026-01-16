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
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var doc = context.Document;
        var shapes = WordImageHelper.GetAllImages(doc, sectionIndex);

        if (shapes.Count == 0)
            return CreateEmptyResult(sectionIndex);

        var imageList = BuildImageList(shapes);

        var result = new
        {
            count = shapes.Count,
            sectionIndex = sectionIndex == -1 ? (int?)null : sectionIndex,
            images = imageList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }

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

    private static List<object> BuildImageList(List<WordShape> shapes)
    {
        List<object> imageList = [];
        for (var i = 0; i < shapes.Count; i++) imageList.Add(BuildImageInfo(shapes[i], i));

        return imageList;
    }

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

    private static (string? alignment, object? position, string? contextText) GetPositionAndContext(WordShape shape)
    {
        if (shape.IsInline)
            return GetInlinePositionAndContext(shape);

        return GetFloatingPositionAndContext(shape);
    }

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

    private static string? TruncateText(string text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        return text.Length > 30 ? text[..30] + "..." : text;
    }

    private static void AddImageDataInfo(Dictionary<string, object?> imageInfo, WordShape shape)
    {
        if (shape.ImageData == null) return;

        imageInfo["imageType"] = shape.ImageData.ImageType.ToString();
        var imageSize = shape.ImageData.ImageSize;
        imageInfo["originalSize"] = new { widthPixels = imageSize.WidthPixels, heightPixels = imageSize.HeightPixels };
    }

    private static void AddOptionalProperties(Dictionary<string, object?> imageInfo, WordShape shape)
    {
        if (!string.IsNullOrEmpty(shape.HRef)) imageInfo["hyperlink"] = shape.HRef;
        if (!string.IsNullOrEmpty(shape.AlternativeText)) imageInfo["altText"] = shape.AlternativeText;
        if (!string.IsNullOrEmpty(shape.Title)) imageInfo["title"] = shape.Title;
    }
}
