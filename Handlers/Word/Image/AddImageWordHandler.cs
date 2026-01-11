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
        var imagePath = parameters.GetOptional<string?>("imagePath");
        var width = parameters.GetOptional<double?>("width");
        var height = parameters.GetOptional<double?>("height");
        var alignment = parameters.GetOptional("alignment", "left");
        var textWrapping = parameters.GetOptional("textWrapping", "inline");
        var caption = parameters.GetOptional<string?>("caption");
        var captionPosition = parameters.GetOptional("captionPosition", "below");
        var linkUrl = parameters.GetOptional<string?>("linkUrl");
        var alternativeText = parameters.GetOptional<string?>("alternativeText");
        var title = parameters.GetOptional<string?>("title");

        if (string.IsNullOrEmpty(imagePath) || !IOFile.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(caption) && captionPosition == "above")
            WordImageHelper.InsertCaption(builder, caption, alignment);

        WordShape shape;
        if (textWrapping == "inline")
        {
            // For inline images, alignment is controlled by paragraph alignment
            var paraAlignment = WordImageHelper.GetAlignment(alignment);
            builder.ParagraphFormat.Alignment = paraAlignment;
            shape = builder.InsertImage(imagePath);

            if (width.HasValue)
                shape.Width = width.Value;

            if (height.HasValue)
                shape.Height = height.Value;

            var currentPara = builder.CurrentParagraph;
            if (currentPara != null)
            {
                currentPara.ParagraphFormat.Alignment = paraAlignment;
                currentPara.ParagraphFormat.ClearFormatting();
                currentPara.ParagraphFormat.Alignment = paraAlignment;
            }

            builder.ParagraphFormat.Alignment = paraAlignment;
        }
        else
        {
            // For floating images, use shape positioning with relative alignment
            shape = builder.InsertImage(imagePath);
            shape.WrapType = WordImageHelper.GetWrapType(textWrapping);

            if (width.HasValue)
                shape.Width = width.Value;

            if (height.HasValue)
                shape.Height = height.Value;

            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
            if (alignment == "center")
                shape.HorizontalAlignment = HorizontalAlignment.Center;
            else if (alignment == "right")
                shape.HorizontalAlignment = HorizontalAlignment.Right;
            else
                shape.HorizontalAlignment = HorizontalAlignment.Left;
        }

        if (!string.IsNullOrEmpty(linkUrl))
            shape.HRef = linkUrl;

        if (!string.IsNullOrEmpty(alternativeText))
            shape.AlternativeText = alternativeText;

        if (!string.IsNullOrEmpty(title))
            shape.Title = title;
        if (!string.IsNullOrEmpty(caption) && captionPosition == "below")
        {
            builder.Writeln(); // New line after image
            WordImageHelper.InsertCaption(builder, caption, alignment);
        }

        if (textWrapping != "inline")
        {
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        }
        else
        {
            // For inline images, ensure the paragraph alignment is preserved
            var currentPara = builder.CurrentParagraph;
            if (currentPara != null)
            {
                var paraAlignment = WordImageHelper.GetAlignment(alignment);
                currentPara.ParagraphFormat.Alignment = paraAlignment;
            }
        }

        MarkModified(context);

        var result = "Image added successfully\n";
        result += $"Image: {Path.GetFileName(imagePath)}\n";
        if (width.HasValue || height.HasValue)
            result +=
                $"Size: {(width.HasValue ? width.Value.ToString(CultureInfo.InvariantCulture) : "auto")} x {(height.HasValue ? height.Value.ToString(CultureInfo.InvariantCulture) : "auto")} pt\n";
        result += $"Alignment: {alignment}\n";
        result += $"Text wrapping: {textWrapping}\n";
        if (!string.IsNullOrEmpty(linkUrl)) result += $"Hyperlink: {linkUrl}\n";
        if (!string.IsNullOrEmpty(alternativeText)) result += $"Alt text: {alternativeText}\n";
        if (!string.IsNullOrEmpty(title)) result += $"Title: {title}\n";
        if (!string.IsNullOrEmpty(caption)) result += $"Caption: {caption} ({captionPosition})\n";

        return result.TrimEnd();
    }
}
