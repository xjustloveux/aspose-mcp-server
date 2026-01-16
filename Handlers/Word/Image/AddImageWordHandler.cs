using System.Globalization;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using IOFile = System.IO.File;
using WordShape = Aspose.Words.Drawing.Shape;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Handler for adding images to Word documents.
/// </summary>
public class AddImageWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds an image to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: imagePath
    ///     Optional: width, height, alignment, textWrapping, caption, captionPosition, linkUrl, alternativeText, title
    /// </param>
    /// <returns>Success message with image details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var imageParams = ExtractImageParameters(parameters);
        ValidateImagePath(imageParams.ImagePath);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        InsertCaptionIfAbove(builder, imageParams);
        var shape = InsertImage(builder, imageParams);
        ApplyShapeProperties(shape, imageParams);
        InsertCaptionIfBelow(builder, imageParams);
        FinalizeAlignment(builder, imageParams);

        MarkModified(context);
        return BuildResultMessage(imageParams);
    }

    /// <summary>
    ///     Extracts image parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted image parameters.</returns>
    private static ImageParameters ExtractImageParameters(OperationParameters parameters)
    {
        return new ImageParameters(
            parameters.GetOptional<string?>("imagePath"),
            parameters.GetOptional<double?>("width"),
            parameters.GetOptional<double?>("height"),
            parameters.GetOptional("alignment", "left"),
            parameters.GetOptional("textWrapping", "inline"),
            parameters.GetOptional<string?>("caption"),
            parameters.GetOptional("captionPosition", "below"),
            parameters.GetOptional<string?>("linkUrl"),
            parameters.GetOptional<string?>("alternativeText"),
            parameters.GetOptional<string?>("title")
        );
    }

    /// <summary>
    ///     Validates that the image path exists.
    /// </summary>
    /// <param name="imagePath">The image path to validate.</param>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    private static void ValidateImagePath(string? imagePath)
    {
        if (string.IsNullOrEmpty(imagePath) || !IOFile.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");
    }

    /// <summary>
    ///     Inserts a caption above the image if configured.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="p">The image parameters.</param>
    private static void InsertCaptionIfAbove(DocumentBuilder builder, ImageParameters p)
    {
        if (!string.IsNullOrEmpty(p.Caption) && p.CaptionPosition == "above")
            WordImageHelper.InsertCaption(builder, p.Caption, p.Alignment);
    }

    /// <summary>
    ///     Inserts the image into the document.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="p">The image parameters.</param>
    /// <returns>The inserted shape.</returns>
    private static WordShape InsertImage(DocumentBuilder builder, ImageParameters p)
    {
        return p.TextWrapping == "inline"
            ? InsertInlineImage(builder, p)
            : InsertFloatingImage(builder, p);
    }

    /// <summary>
    ///     Inserts an inline image.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="p">The image parameters.</param>
    /// <returns>The inserted shape.</returns>
    private static WordShape InsertInlineImage(DocumentBuilder builder, ImageParameters p)
    {
        var paraAlignment = WordImageHelper.GetAlignment(p.Alignment);
        builder.ParagraphFormat.Alignment = paraAlignment;
        var shape = builder.InsertImage(p.ImagePath);

        if (p.Width.HasValue) shape.Width = p.Width.Value;
        if (p.Height.HasValue) shape.Height = p.Height.Value;

        var currentPara = builder.CurrentParagraph;
        if (currentPara != null)
        {
            currentPara.ParagraphFormat.Alignment = paraAlignment;
            currentPara.ParagraphFormat.ClearFormatting();
            currentPara.ParagraphFormat.Alignment = paraAlignment;
        }

        builder.ParagraphFormat.Alignment = paraAlignment;
        return shape;
    }

    /// <summary>
    ///     Inserts a floating image with text wrapping.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="p">The image parameters.</param>
    /// <returns>The inserted shape.</returns>
    private static WordShape InsertFloatingImage(DocumentBuilder builder, ImageParameters p)
    {
        var shape = builder.InsertImage(p.ImagePath);
        shape.WrapType = WordImageHelper.GetWrapType(p.TextWrapping);

        if (p.Width.HasValue) shape.Width = p.Width.Value;
        if (p.Height.HasValue) shape.Height = p.Height.Value;

        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        shape.HorizontalAlignment = p.Alignment switch
        {
            "center" => HorizontalAlignment.Center,
            "right" => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Left
        };
        return shape;
    }

    /// <summary>
    ///     Applies additional properties to the shape.
    /// </summary>
    /// <param name="shape">The shape to configure.</param>
    /// <param name="p">The image parameters.</param>
    private static void ApplyShapeProperties(WordShape shape, ImageParameters p)
    {
        if (!string.IsNullOrEmpty(p.LinkUrl)) shape.HRef = p.LinkUrl;
        if (!string.IsNullOrEmpty(p.AlternativeText)) shape.AlternativeText = p.AlternativeText;
        if (!string.IsNullOrEmpty(p.Title)) shape.Title = p.Title;
    }

    /// <summary>
    ///     Inserts a caption below the image if configured.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="p">The image parameters.</param>
    private static void InsertCaptionIfBelow(DocumentBuilder builder, ImageParameters p)
    {
        if (!string.IsNullOrEmpty(p.Caption) && p.CaptionPosition == "below")
        {
            builder.Writeln();
            WordImageHelper.InsertCaption(builder, p.Caption, p.Alignment);
        }
    }

    /// <summary>
    ///     Finalizes the paragraph alignment after image insertion.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="p">The image parameters.</param>
    private static void FinalizeAlignment(DocumentBuilder builder, ImageParameters p)
    {
        if (p.TextWrapping != "inline")
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        else if (builder.CurrentParagraph != null)
            builder.CurrentParagraph.ParagraphFormat.Alignment = WordImageHelper.GetAlignment(p.Alignment);
    }

    /// <summary>
    ///     Builds the result message for a successful image addition.
    /// </summary>
    /// <param name="p">The image parameters.</param>
    /// <returns>The formatted result message.</returns>
    private static string BuildResultMessage(ImageParameters p)
    {
        var result = $"Image added successfully\nImage: {Path.GetFileName(p.ImagePath)}\n";
        if (p.Width.HasValue || p.Height.HasValue)
            result += $"Size: {(p.Width.HasValue ? p.Width.Value.ToString(CultureInfo.InvariantCulture) : "auto")} x " +
                      $"{(p.Height.HasValue ? p.Height.Value.ToString(CultureInfo.InvariantCulture) : "auto")} pt\n";
        result += $"Alignment: {p.Alignment}\nText wrapping: {p.TextWrapping}\n";
        if (!string.IsNullOrEmpty(p.LinkUrl)) result += $"Hyperlink: {p.LinkUrl}\n";
        if (!string.IsNullOrEmpty(p.AlternativeText)) result += $"Alt text: {p.AlternativeText}\n";
        if (!string.IsNullOrEmpty(p.Title)) result += $"Title: {p.Title}\n";
        if (!string.IsNullOrEmpty(p.Caption)) result += $"Caption: {p.Caption} ({p.CaptionPosition})\n";
        return result.TrimEnd();
    }

    /// <summary>
    ///     Record to hold image insertion parameters.
    /// </summary>
    private record ImageParameters(
        string? ImagePath,
        double? Width,
        double? Height,
        string Alignment,
        string TextWrapping,
        string? Caption,
        string CaptionPosition,
        string? LinkUrl,
        string? AlternativeText,
        string? Title);
}
