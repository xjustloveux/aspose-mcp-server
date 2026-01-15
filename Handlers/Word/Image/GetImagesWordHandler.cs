using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

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

        List<object> imageList = [];
        for (var i = 0; i < shapes.Count; i++)
        {
            var shape = shapes[i];
            string? contextText = null;
            string? alignment = null;
            object? position = null;

            if (shape.IsInline)
            {
                if (shape.ParentNode is WordParagraph parentPara)
                {
                    alignment = parentPara.ParagraphFormat.Alignment.ToString();
                    var paraText = parentPara.GetText().Trim();
                    if (paraText.Length > 30) paraText = paraText[..30] + "...";
                    if (!string.IsNullOrEmpty(paraText)) contextText = paraText;
                }
                else
                {
                    position = new { x = shape.Left, y = shape.Top };
                }
            }
            else
            {
                position = new
                {
                    x = Math.Round(shape.Left, 1),
                    y = Math.Round(shape.Top, 1),
                    horizontalAlignment = shape.HorizontalAlignment.ToString(),
                    verticalAlignment = shape.VerticalAlignment.ToString(),
                    wrapType = shape.WrapType.ToString()
                };
                if (shape.GetAncestor(NodeType.Paragraph) is WordParagraph nearestPara)
                {
                    var paraText = nearestPara.GetText().Trim();
                    if (paraText.Length > 30) paraText = paraText[..30] + "...";
                    if (!string.IsNullOrEmpty(paraText)) contextText = paraText;
                }
            }

            var imageInfo = new Dictionary<string, object?>
            {
                ["index"] = i,
                ["name"] = string.IsNullOrEmpty(shape.Name) ? null : shape.Name,
                ["width"] = shape.Width,
                ["height"] = shape.Height,
                ["isInline"] = shape.IsInline
            };

            if (alignment != null) imageInfo["alignment"] = alignment;
            if (position != null) imageInfo["position"] = position;
            if (contextText != null) imageInfo["context"] = contextText;

            if (shape.ImageData != null)
            {
                imageInfo["imageType"] = shape.ImageData.ImageType.ToString();
                var imageSize = shape.ImageData.ImageSize;
                imageInfo["originalSize"] = new
                    { widthPixels = imageSize.WidthPixels, heightPixels = imageSize.HeightPixels };
            }

            if (!string.IsNullOrEmpty(shape.HRef)) imageInfo["hyperlink"] = shape.HRef;
            if (!string.IsNullOrEmpty(shape.AlternativeText)) imageInfo["altText"] = shape.AlternativeText;
            if (!string.IsNullOrEmpty(shape.Title)) imageInfo["title"] = shape.Title;

            imageList.Add(imageInfo);
        }

        var result = new
        {
            count = shapes.Count,
            sectionIndex = sectionIndex == -1 ? (int?)null : sectionIndex,
            images = imageList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }
}
