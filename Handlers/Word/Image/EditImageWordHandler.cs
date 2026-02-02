using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for editing images in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractImageParameters(parameters);
        var doc = context.Document;

        var allImages = WordImageHelper.GetAllImages(doc, p.SectionIndex);
        ValidateImageIndex(p.ImageIndex, allImages.Count);

        var shape = allImages[p.ImageIndex];

        ApplySizeProperties(shape, p);
        ApplyAlignmentProperties(shape, p);
        ApplyTextWrappingProperties(shape, p);
        ApplyMetadataProperties(shape, p);

        MarkModified(context);
        return new SuccessResult { Message = $"Image {p.ImageIndex} edited ({BuildChangesDescription(p)})" };
    }

    /// <summary>
    ///     Extracts image edit parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted image edit parameters.</returns>
    private static ImageEditParameters ExtractImageParameters(OperationParameters parameters)
    {
        return new ImageEditParameters(
            parameters.GetOptional("imageIndex", 0),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional<double?>("width"),
            parameters.GetOptional<double?>("height"),
            parameters.GetOptional<string?>("alignment"),
            parameters.GetOptional<string?>("textWrapping"),
            parameters.GetOptional<bool?>("aspectRatioLocked"),
            parameters.GetOptional<string?>("horizontalAlignment"),
            parameters.GetOptional<string?>("verticalAlignment"),
            parameters.GetOptional<string?>("alternativeText"),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<string?>("linkUrl")
        );
    }

    /// <summary>
    ///     Validates that the image index is within range.
    /// </summary>
    /// <param name="imageIndex">The image index to validate.</param>
    /// <param name="count">The total number of images.</param>
    /// <exception cref="ArgumentException">Thrown when image index is out of range.</exception>
    private static void ValidateImageIndex(int imageIndex, int count)
    {
        if (imageIndex < 0 || imageIndex >= count)
            throw new ArgumentException($"Image index {imageIndex} is out of range (document has {count} images)");
    }

    /// <summary>
    ///     Applies size properties to the shape.
    /// </summary>
    /// <param name="shape">The shape to configure.</param>
    /// <param name="p">The image edit parameters.</param>
    private static void ApplySizeProperties(WordShape shape, ImageEditParameters p)
    {
        if (p is { Width: not null, Height: not null })
            shape.AspectRatioLocked = false;

        if (p.Width.HasValue) shape.Width = p.Width.Value;
        if (p.Height.HasValue) shape.Height = p.Height.Value;

        if (p.AspectRatioLocked.HasValue) shape.AspectRatioLocked = p.AspectRatioLocked.Value;
    }

    /// <summary>
    ///     Applies alignment properties to the shape.
    /// </summary>
    /// <param name="shape">The shape to configure.</param>
    /// <param name="p">The image edit parameters.</param>
    private static void ApplyAlignmentProperties(WordShape shape, ImageEditParameters p)
    {
        var alignmentValue = p.Alignment ?? "left";
        if (!string.IsNullOrEmpty(alignmentValue) && shape.ParentNode is WordParagraph parentPara)
            parentPara.ParagraphFormat.Alignment = WordImageHelper.GetAlignment(alignmentValue);
    }

    /// <summary>
    ///     Applies text wrapping properties to the shape.
    /// </summary>
    /// <param name="shape">The shape to configure.</param>
    /// <param name="p">The image edit parameters.</param>
    private static void ApplyTextWrappingProperties(WordShape shape, ImageEditParameters p)
    {
        var textWrappingValue = p.TextWrapping ?? "inline";
        if (string.IsNullOrEmpty(textWrappingValue) && shape.WrapType == WrapType.Inline) return;

        if (!string.IsNullOrEmpty(textWrappingValue))
            shape.WrapType = WordImageHelper.GetWrapType(textWrappingValue);

        if (textWrappingValue != "inline" && shape.WrapType != WrapType.Inline)
            ApplyFloatingPositionProperties(shape, p);
    }

    /// <summary>
    ///     Applies floating position properties to the shape.
    /// </summary>
    /// <param name="shape">The shape to configure.</param>
    /// <param name="p">The image edit parameters.</param>
    private static void ApplyFloatingPositionProperties(WordShape shape, ImageEditParameters p)
    {
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

        var hAlign = p.HorizontalAlignment ?? "left";
        if (!string.IsNullOrEmpty(hAlign))
            shape.HorizontalAlignment = WordImageHelper.GetHorizontalAlignment(hAlign);

        var vAlign = p.VerticalAlignment ?? "top";
        if (!string.IsNullOrEmpty(vAlign))
            shape.VerticalAlignment = WordImageHelper.GetVerticalAlignment(vAlign);
    }

    /// <summary>
    ///     Applies metadata properties to the shape.
    /// </summary>
    /// <param name="shape">The shape to configure.</param>
    /// <param name="p">The image edit parameters.</param>
    private static void ApplyMetadataProperties(WordShape shape, ImageEditParameters p)
    {
        if (!string.IsNullOrEmpty(p.AlternativeText)) shape.AlternativeText = p.AlternativeText;
        if (!string.IsNullOrEmpty(p.Title)) shape.Title = p.Title;
        if (p.LinkUrl != null) shape.HRef = p.LinkUrl;
    }

    /// <summary>
    ///     Builds a description of the changes made.
    /// </summary>
    /// <param name="p">The image edit parameters.</param>
    /// <returns>A description of the changes.</returns>
    private static string BuildChangesDescription(ImageEditParameters p)
    {
        List<string> changes = [];
        if (p.Width.HasValue) changes.Add($"Width: {p.Width.Value}");
        if (p.Height.HasValue) changes.Add($"Height: {p.Height.Value}");
        if (p.Alignment != null) changes.Add($"Alignment: {p.Alignment}");
        if (p.TextWrapping != null) changes.Add($"Text wrapping: {p.TextWrapping}");
        if (p.LinkUrl != null)
            changes.Add(string.IsNullOrEmpty(p.LinkUrl) ? "Hyperlink: removed" : $"Hyperlink: {p.LinkUrl}");
        if (p.AlternativeText != null) changes.Add($"Alt text: {p.AlternativeText}");
        if (p.Title != null) changes.Add($"Title: {p.Title}");

        return changes.Count > 0 ? string.Join(", ", changes) : "properties";
    }

    /// <summary>
    ///     Record to hold image edit parameters.
    /// </summary>
    private sealed record ImageEditParameters(
        int ImageIndex,
        int SectionIndex,
        double? Width,
        double? Height,
        string? Alignment,
        string? TextWrapping,
        bool? AspectRatioLocked,
        string? HorizontalAlignment,
        string? VerticalAlignment,
        string? AlternativeText,
        string? Title,
        string? LinkUrl);
}
