using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for editing images in Word documents.
/// </summary>
public class EditImageWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits image properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: imageIndex
    ///     Optional: sectionIndex, width, height, alignment, textWrapping, aspectRatioLocked,
    ///     horizontalAlignment, verticalAlignment, alternativeText, title, linkUrl
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var imageIndex = parameters.GetOptional("imageIndex", 0);
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var width = parameters.GetOptional<double?>("width");
        var height = parameters.GetOptional<double?>("height");
        var alignment = parameters.GetOptional<string?>("alignment");
        var textWrapping = parameters.GetOptional<string?>("textWrapping");
        var aspectRatioLocked = parameters.GetOptional<bool?>("aspectRatioLocked");
        var horizontalAlignment = parameters.GetOptional<string?>("horizontalAlignment");
        var verticalAlignment = parameters.GetOptional<string?>("verticalAlignment");
        var alternativeText = parameters.GetOptional<string?>("alternativeText");
        var title = parameters.GetOptional<string?>("title");
        var linkUrl = parameters.GetOptional<string?>("linkUrl");

        var doc = context.Document;

        var allImages = WordImageHelper.GetAllImages(doc, sectionIndex);

        if (imageIndex < 0 || imageIndex >= allImages.Count)
            throw new ArgumentException(
                $"Image index {imageIndex} is out of range (document has {allImages.Count} images)");

        var shape = allImages[imageIndex];

        // Apply size properties
        if (width.HasValue)
            shape.Width = width.Value;

        if (height.HasValue)
            shape.Height = height.Value;

        if (aspectRatioLocked.HasValue)
            shape.AspectRatioLocked = aspectRatioLocked.Value;

        var alignmentValue = alignment ?? "left";
        if (!string.IsNullOrEmpty(alignmentValue) && shape.ParentNode is WordParagraph parentPara)
            parentPara.ParagraphFormat.Alignment = WordImageHelper.GetAlignment(alignmentValue);

        var textWrappingValue = textWrapping ?? "inline";
        if (!string.IsNullOrEmpty(textWrappingValue))
        {
            shape.WrapType = WordImageHelper.GetWrapType(textWrappingValue);

            if (textWrappingValue != "inline")
            {
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
                shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

                var hAlign = horizontalAlignment ?? "left";
                if (!string.IsNullOrEmpty(hAlign))
                    shape.HorizontalAlignment = WordImageHelper.GetHorizontalAlignment(hAlign);

                var vAlign = verticalAlignment ?? "top";
                if (!string.IsNullOrEmpty(vAlign))
                    shape.VerticalAlignment = WordImageHelper.GetVerticalAlignment(vAlign);
            }
        }
        else if (shape.WrapType != WrapType.Inline)
        {
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

            var hAlign = horizontalAlignment ?? "left";
            if (!string.IsNullOrEmpty(hAlign))
                shape.HorizontalAlignment = WordImageHelper.GetHorizontalAlignment(hAlign);

            var vAlign = verticalAlignment ?? "top";
            if (!string.IsNullOrEmpty(vAlign)) shape.VerticalAlignment = WordImageHelper.GetVerticalAlignment(vAlign);
        }

        if (!string.IsNullOrEmpty(alternativeText))
            shape.AlternativeText = alternativeText;

        if (!string.IsNullOrEmpty(title))
            shape.Title = title;

        // HRef property doesn't accept null, use empty string to clear
        if (linkUrl != null)
            shape.HRef = linkUrl;

        MarkModified(context);

        List<string> changes = [];
        if (width.HasValue) changes.Add($"Width: {width.Value}");
        if (height.HasValue) changes.Add($"Height: {height.Value}");
        if (alignment != null) changes.Add($"Alignment: {alignment}");
        if (textWrapping != null) changes.Add($"Text wrapping: {textWrapping}");
        if (linkUrl != null)
            changes.Add(string.IsNullOrEmpty(linkUrl) ? "Hyperlink: removed" : $"Hyperlink: {linkUrl}");
        if (alternativeText != null) changes.Add($"Alt text: {alternativeText}");
        if (title != null) changes.Add($"Title: {title}");

        var changesDesc = changes.Count > 0 ? string.Join(", ", changes) : "properties";

        return Success($"Image {imageIndex} edited ({changesDesc})");
    }
}
