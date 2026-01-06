using System.ComponentModel;
using System.Globalization;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word images (add, edit, delete, get, replace, extract)
///     Merges: WordAddImageTool, WordEditImageTool, WordDeleteImageTool, WordGetImagesTool, WordReplaceImageTool,
///     WordExtractImagesTool
/// </summary>
[McpServerToolType]
public class WordImageTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordImageTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordImageTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word image operation (add, edit, delete, get, replace, extract).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, replace, extract.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="outputDir">Output directory (for extract).</param>
    /// <param name="imagePath">Image file path (for add/replace).</param>
    /// <param name="imageIndex">Image index (0-based, for edit/delete/replace).</param>
    /// <param name="sectionIndex">Section index (0-based, use -1 for all sections).</param>
    /// <param name="width">Image width in points.</param>
    /// <param name="height">Image height in points.</param>
    /// <param name="alignment">Horizontal alignment: left, center, right.</param>
    /// <param name="textWrapping">Text wrapping: inline, square, tight, through, topAndBottom, none.</param>
    /// <param name="caption">Image caption text (for add).</param>
    /// <param name="captionPosition">Caption position: above, below (for add).</param>
    /// <param name="aspectRatioLocked">Lock aspect ratio (for edit).</param>
    /// <param name="horizontalAlignment">Horizontal alignment for floating images (for edit).</param>
    /// <param name="verticalAlignment">Vertical alignment for floating images (for edit).</param>
    /// <param name="alternativeText">Alternative text for accessibility.</param>
    /// <param name="title">Image title.</param>
    /// <param name="linkUrl">Hyperlink URL for the image.</param>
    /// <param name="newImagePath">New image file path for replace operation.</param>
    /// <param name="preserveSize">Preserve original image size for replace operation.</param>
    /// <param name="smartFit">Smart fit to avoid distortion for replace operation.</param>
    /// <param name="preservePosition">Preserve original image position for replace operation.</param>
    /// <param name="prefix">Filename prefix for extracted images.</param>
    /// <param name="extractImageIndex">Specific image index to extract.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_image")]
    [Description(@"Manage Word document images. Supports 6 operations: add, edit, delete, get, replace, extract.

Usage examples:
- Add image: word_image(operation='add', path='doc.docx', imagePath='image.png', width=200)
- Edit image: word_image(operation='edit', path='doc.docx', imageIndex=0, width=300, height=200)
- Delete image: word_image(operation='delete', path='doc.docx', imageIndex=0)
- Get all images: word_image(operation='get', path='doc.docx')
- Replace image: word_image(operation='replace', path='doc.docx', imageIndex=0, imagePath='new_image.png')
- Extract images: word_image(operation='extract', path='doc.docx', outputDir='images/')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a new image (required params: path, imagePath)
- 'edit': Edit existing image (required params: path, imageIndex)
- 'delete': Delete an image (required params: path, imageIndex)
- 'get': Get all images info (required params: path)
- 'replace': Replace an image (required params: path, imageIndex, imagePath)
- 'extract': Extract all images (required params: path, outputDir)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Output directory (required for extract operation)")]
        string? outputDir = null,
        [Description("Image file path (required for add and replace operations)")]
        string? imagePath = null,
        [Description(
            "Image index (0-based, required for edit, delete, and replace operations). Note: After delete operations, subsequent image indices will shift automatically. Use 'get' operation to refresh indices.")]
        int? imageIndex = null,
        [Description("Section index (0-based, optional, default: 0, use -1 to search all sections)")]
        int sectionIndex = 0,
        [Description("Image width in points (72 pts = 1 inch, optional, for add/edit operations)")]
        double? width = null,
        [Description("Image height in points (72 pts = 1 inch, optional, for add/edit operations)")]
        double? height = null,
        [Description("Horizontal alignment: left, center, right (optional, for add/edit operations)")]
        string alignment = "left",
        [Description(
            "Text wrapping: inline, square, tight, through, topAndBottom, none (optional, for add/edit operations)")]
        string textWrapping = "inline",
        [Description("Image caption text (optional, for add operation)")]
        string? caption = null,
        [Description("Caption position: above, below (optional, for add operation)")]
        string captionPosition = "below",
        [Description("Lock aspect ratio (optional, for edit operation)")]
        bool? aspectRatioLocked = null,
        [Description("Horizontal alignment for floating images: left, center, right (optional, for edit operation)")]
        string? horizontalAlignment = null,
        [Description("Vertical alignment for floating images: top, center, bottom (optional, for edit operation)")]
        string? verticalAlignment = null,
        [Description("Alternative text for accessibility (optional, for add/edit operation)")]
        string? alternativeText = null,
        [Description("Image title (optional, for add/edit operation)")]
        string? title = null,
        [Description(
            "Hyperlink URL for the image. When clicked, opens the specified URL (optional, for add/edit operation). Use empty string to remove existing hyperlink.")]
        string? linkUrl = null,
        [Description("New image file path (required for replace operation)")]
        string? newImagePath = null,
        [Description("Preserve original image size (default: true, for replace operation)")]
        bool preserveSize = true,
        [Description(
            "When true, keeps original width and calculates height proportionally based on new image aspect ratio (avoids distortion when aspect ratios differ, default: false, for replace operation). Only applies when preserveSize is true.")]
        bool smartFit = false,
        [Description("Preserve original image position and wrapping (default: true, for replace operation)")]
        bool preservePosition = true,
        [Description("Filename prefix for extracted images (optional, default: 'image', for extract operation)")]
        string prefix = "image",
        [Description(
            "Specific image index to extract (0-based, optional, for extract operation). If not provided, extracts all images.")]
        int? extractImageIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddImage(ctx, outputPath, imagePath!, width, height, alignment, textWrapping, caption,
                captionPosition, linkUrl, alternativeText, title),
            "edit" => EditImage(ctx, outputPath, imageIndex ?? 0, sectionIndex, width, height, alignment, textWrapping,
                aspectRatioLocked, horizontalAlignment, verticalAlignment, alternativeText, title, linkUrl),
            "delete" => DeleteImage(ctx, outputPath, imageIndex ?? 0, sectionIndex),
            "get" => GetImages(ctx, sectionIndex),
            "replace" => ReplaceImage(ctx, outputPath, imageIndex ?? 0, newImagePath ?? imagePath!, preserveSize,
                smartFit, preservePosition, sectionIndex),
            "extract" => ExtractImages(ctx, outputDir!, prefix, extractImageIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="imagePath">The image file path.</param>
    /// <param name="width">The image width in points.</param>
    /// <param name="height">The image height in points.</param>
    /// <param name="alignment">The horizontal alignment (left, center, right).</param>
    /// <param name="textWrapping">The text wrapping style.</param>
    /// <param name="caption">The caption text.</param>
    /// <param name="captionPosition">The caption position (above, below).</param>
    /// <param name="linkUrl">The hyperlink URL for the image.</param>
    /// <param name="alternativeText">The alternative text for accessibility.</param>
    /// <param name="title">The image title.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    private static string AddImage(DocumentContext<Document> ctx, string? outputPath, string imagePath,
        double? width, double? height, string alignment, string textWrapping,
        string? caption, string captionPosition, string? linkUrl, string? alternativeText, string? title)
    {
        if (!File.Exists(imagePath)) throw new FileNotFoundException($"Image file not found: {imagePath}");

        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(caption) && captionPosition == "above")
            InsertCaption(builder, caption, alignment);

        Shape shape;
        if (textWrapping == "inline")
        {
            // For inline images, alignment is controlled by paragraph alignment
            var paraAlignment = GetAlignment(alignment);
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
            shape.WrapType = GetWrapType(textWrapping);

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
            InsertCaption(builder, caption, alignment);
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
                var paraAlignment = GetAlignment(alignment);
                currentPara.ParagraphFormat.Alignment = paraAlignment;
            }
        }

        ctx.Save(outputPath);

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
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Edits image properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="imageIndex">The zero-based image index.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <param name="width">The image width in points.</param>
    /// <param name="height">The image height in points.</param>
    /// <param name="alignment">The horizontal alignment.</param>
    /// <param name="textWrapping">The text wrapping style.</param>
    /// <param name="aspectRatioLocked">Whether to lock the aspect ratio.</param>
    /// <param name="horizontalAlignment">The horizontal alignment for floating images.</param>
    /// <param name="verticalAlignment">The vertical alignment for floating images.</param>
    /// <param name="alternativeText">The alternative text for accessibility.</param>
    /// <param name="title">The image title.</param>
    /// <param name="linkUrl">The hyperlink URL for the image.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the image index is out of range.</exception>
    private static string EditImage(DocumentContext<Document> ctx, string? outputPath, int imageIndex, int sectionIndex,
        double? width, double? height, string? alignment, string? textWrapping,
        bool? aspectRatioLocked, string? horizontalAlignment, string? verticalAlignment,
        string? alternativeText, string? title, string? linkUrl)
    {
        var doc = ctx.Document;

        var allImages = GetAllImages(doc, sectionIndex);

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
        if (!string.IsNullOrEmpty(alignmentValue))
            if (shape.ParentNode is Paragraph parentPara)
                parentPara.ParagraphFormat.Alignment = GetAlignment(alignmentValue);

        var textWrappingValue = textWrapping ?? "inline";
        if (!string.IsNullOrEmpty(textWrappingValue))
        {
            shape.WrapType = GetWrapType(textWrappingValue);

            if (textWrappingValue != "inline")
            {
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
                shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

                var hAlign = horizontalAlignment ?? "left";
                if (!string.IsNullOrEmpty(hAlign)) shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);

                var vAlign = verticalAlignment ?? "top";
                if (!string.IsNullOrEmpty(vAlign)) shape.VerticalAlignment = GetVerticalAlignment(vAlign);
            }
        }
        else if (shape.WrapType != WrapType.Inline)
        {
            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

            var hAlign = horizontalAlignment ?? "left";
            if (!string.IsNullOrEmpty(hAlign)) shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);

            var vAlign = verticalAlignment ?? "top";
            if (!string.IsNullOrEmpty(vAlign)) shape.VerticalAlignment = GetVerticalAlignment(vAlign);
        }

        if (!string.IsNullOrEmpty(alternativeText))
            shape.AlternativeText = alternativeText;

        if (!string.IsNullOrEmpty(title))
            shape.Title = title;

        // HRef property doesn't accept null, use empty string to clear
        if (linkUrl != null)
            shape.HRef = linkUrl;

        ctx.Save(outputPath);

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

        var result = $"Image {imageIndex} edited successfully ({changesDesc})\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes an image from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="imageIndex">The zero-based image index.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the image index is out of range.</exception>
    private static string DeleteImage(DocumentContext<Document> ctx, string? outputPath, int imageIndex,
        int sectionIndex)
    {
        var doc = ctx.Document;

        var allImages = GetAllImages(doc, sectionIndex);

        if (imageIndex < 0 || imageIndex >= allImages.Count)
            throw new ArgumentException(
                $"Image index {imageIndex} is out of range (document has {allImages.Count} images)");

        var shapeToDelete = allImages[imageIndex];

        var imageInfo = $"Image #{imageIndex}";
        if (shapeToDelete.HasImage)
            try
            {
                imageInfo += $" (Width: {shapeToDelete.Width:F1} pt, Height: {shapeToDelete.Height:F1} pt)";
            }
            catch (Exception ex)
            {
                // Size information may not be available, but this is not critical
                Console.Error.WriteLine($"[WARN] Failed to get image size information: {ex.Message}");
                // Continue without the size information
            }

        shapeToDelete.Remove();

        ctx.Save(outputPath);

        var remainingCount = GetAllImages(doc, sectionIndex).Count;

        var result = $"{imageInfo} deleted successfully\n";
        result += $"Remaining images in document: {remainingCount}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Gets all images from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <returns>A JSON string containing image information.</returns>
    private static string GetImages(DocumentContext<Document> ctx, int sectionIndex)
    {
        var doc = ctx.Document;

        var shapes = GetAllImages(doc, sectionIndex);

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
            return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
        }

        List<object> imageList = [];
        for (var i = 0; i < shapes.Count; i++)
        {
            var shape = shapes[i];
            string? context = null;
            string? alignment = null;
            object? position = null;

            if (shape.IsInline)
            {
                if (shape.ParentNode is Paragraph parentPara)
                {
                    alignment = parentPara.ParagraphFormat.Alignment.ToString();
                    var paraText = parentPara.GetText().Trim();
                    if (paraText.Length > 30) paraText = paraText[..30] + "...";
                    if (!string.IsNullOrEmpty(paraText)) context = paraText;
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
                if (shape.GetAncestor(NodeType.Paragraph) is Paragraph nearestPara)
                {
                    var paraText = nearestPara.GetText().Trim();
                    if (paraText.Length > 30) paraText = paraText[..30] + "...";
                    if (!string.IsNullOrEmpty(paraText)) context = paraText;
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
            if (context != null) imageInfo["context"] = context;

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

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Replaces an existing image with a new one.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="imageIndex">The zero-based image index.</param>
    /// <param name="newImagePath">The new image file path.</param>
    /// <param name="preserveSize">Whether to preserve the original image size.</param>
    /// <param name="smartFit">Whether to use smart fit for proportional sizing.</param>
    /// <param name="preservePosition">Whether to preserve the original image position.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the new image file is not found.</exception>
    /// <exception cref="ArgumentException">Thrown when the image index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when an error occurs while replacing the image.</exception>
    private static string ReplaceImage(DocumentContext<Document> ctx, string? outputPath, int imageIndex,
        string newImagePath,
        bool preserveSize, bool smartFit, bool preservePosition, int sectionIndex)
    {
        SecurityHelper.ValidateFilePath(newImagePath, "newImagePath", true);

        if (!File.Exists(newImagePath)) throw new FileNotFoundException($"Image file not found: {newImagePath}");

        var doc = ctx.Document;

        var allImages = GetAllImages(doc, sectionIndex);

        if (imageIndex < 0 || imageIndex >= allImages.Count)
            throw new ArgumentException(
                $"Image index {imageIndex} is out of range (document has {allImages.Count} images)");

        var shapeToReplace = allImages[imageIndex];

        var originalWidth = shapeToReplace.Width;
        var originalHeight = shapeToReplace.Height;
        var originalWrapType = shapeToReplace.WrapType;
        HorizontalAlignment? originalHorizontalAlignment = null;
        VerticalAlignment? originalVerticalAlignment = null;
        RelativeHorizontalPosition? originalRelativeHorizontalPosition = null;
        RelativeVerticalPosition? originalRelativeVerticalPosition = null;
        double? originalLeft = null;
        double? originalTop = null;

        if (preservePosition)
        {
            originalHorizontalAlignment = shapeToReplace.HorizontalAlignment;
            originalVerticalAlignment = shapeToReplace.VerticalAlignment;
            originalRelativeHorizontalPosition = shapeToReplace.RelativeHorizontalPosition;
            originalRelativeVerticalPosition = shapeToReplace.RelativeVerticalPosition;
            originalLeft = shapeToReplace.Left;
            originalTop = shapeToReplace.Top;
        }

        try
        {
            shapeToReplace.ImageData.SetImage(newImagePath);

            if (preserveSize)
            {
                if (smartFit)
                {
                    // Calculate proportional height based on new image's aspect ratio
                    var newImageSize = shapeToReplace.ImageData.ImageSize;
                    if (newImageSize.WidthPixels > 0)
                    {
                        var newAspectRatio = (double)newImageSize.HeightPixels / newImageSize.WidthPixels;
                        shapeToReplace.Width = originalWidth;
                        shapeToReplace.Height = originalWidth * newAspectRatio;
                    }
                    else
                    {
                        // Fallback to original size if aspect ratio can't be calculated
                        shapeToReplace.Width = originalWidth;
                        shapeToReplace.Height = originalHeight;
                    }
                }
                else
                {
                    shapeToReplace.Width = originalWidth;
                    shapeToReplace.Height = originalHeight;
                }
            }

            if (preservePosition)
            {
                shapeToReplace.WrapType = originalWrapType;
                if (originalHorizontalAlignment.HasValue)
                    shapeToReplace.HorizontalAlignment = originalHorizontalAlignment.Value;
                if (originalVerticalAlignment.HasValue)
                    shapeToReplace.VerticalAlignment = originalVerticalAlignment.Value;
                if (originalRelativeHorizontalPosition.HasValue)
                    shapeToReplace.RelativeHorizontalPosition = originalRelativeHorizontalPosition.Value;
                if (originalRelativeVerticalPosition.HasValue)
                    shapeToReplace.RelativeVerticalPosition = originalRelativeVerticalPosition.Value;
                if (originalLeft.HasValue)
                    shapeToReplace.Left = originalLeft.Value;
                if (originalTop.HasValue)
                    shapeToReplace.Top = originalTop.Value;
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error occurred while replacing image: {ex.Message}", ex);
        }

        ctx.Save(outputPath);

        var result = $"Image #{imageIndex} replaced successfully\n";
        result += $"New image: {Path.GetFileName(newImagePath)}\n";
        if (preserveSize)
        {
            if (smartFit)
                result +=
                    $"Smart fit: width preserved ({originalWidth:F1} pt), height calculated proportionally ({shapeToReplace.Height:F1} pt)\n";
            else
                result += $"Preserved size: {originalWidth:F1} pt x {originalHeight:F1} pt\n";
        }

        if (preservePosition) result += "Preserved position and wrapping\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Extracts images from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputDir">The output directory for extracted images.</param>
    /// <param name="prefix">The filename prefix for extracted images.</param>
    /// <param name="extractImageIndex">The specific image index to extract, or null for all images.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the image index is out of range.</exception>
    private static string ExtractImages(DocumentContext<Document> ctx, string outputDir, string prefix,
        int? extractImageIndex)
    {
        SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

        Directory.CreateDirectory(outputDir);

        var doc = ctx.Document;
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();

        if (shapes.Count == 0) return "No images found in document";

        // Validate extractImageIndex if provided
        if (extractImageIndex.HasValue)
            if (extractImageIndex.Value < 0 || extractImageIndex.Value >= shapes.Count)
                throw new ArgumentException(
                    $"Image index {extractImageIndex.Value} is out of range (document has {shapes.Count} images)");

        List<string> extractedFiles = [];

        // Determine which images to extract
        var startIndex = extractImageIndex ?? 0;
        var endIndex = extractImageIndex.HasValue ? extractImageIndex.Value + 1 : shapes.Count;

        for (var i = startIndex; i < endIndex; i++)
        {
            var shape = shapes[i];
            var imageData = shape.ImageData;

            // Use FileFormatUtil for reliable image type detection
            var extension = FileFormatUtil.ImageTypeToExtension(imageData.ImageType);
            if (string.IsNullOrEmpty(extension) || extension == ".")
                extension = ".img";
            // Remove leading dot if present for consistent filename handling
            if (extension.StartsWith('.'))
                extension = extension.Substring(1);

            var safePrefix = SecurityHelper.SanitizeFileName(prefix);
            var filename = $"{safePrefix}_{i + 1:D3}.{extension}";
            var outputFilePath = Path.Combine(outputDir, filename);

            using (var stream = File.Create(outputFilePath))
            {
                imageData.Save(stream);
            }

            extractedFiles.Add(outputFilePath);
        }

        if (extractImageIndex.HasValue)
            return $"Successfully extracted image #{extractImageIndex.Value} to: {outputDir}\n" +
                   $"File: {Path.GetFileName(extractedFiles[0])}";

        return $"Successfully extracted {shapes.Count} images to: {outputDir}\n" +
               $"File list:\n" + string.Join("\n",
                   extractedFiles.Select(f => $"  - {Path.GetFileName(f)}"));
    }

    /// <summary>
    ///     Gets all images from the document or a specific section.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="sectionIndex">The zero-based section index (-1 for all sections).</param>
    /// <returns>A list of Shape objects representing images.</returns>
    /// <exception cref="ArgumentException">Thrown when the section index is out of range.</exception>
    private static List<Shape> GetAllImages(Document doc, int sectionIndex)
    {
        List<Shape> allImages = [];

        if (sectionIndex == -1)
        {
            foreach (var section in doc.Sections.Cast<Section>())
            {
                var shapes = section.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
                allImages.AddRange(shapes);
            }
        }
        else
        {
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException(
                    $"Section index {sectionIndex} is out of range (document has {doc.Sections.Count} sections)");

            var section = doc.Sections[sectionIndex];
            allImages = section.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        }

        return allImages;
    }

    /// <summary>
    ///     Converts alignment string to ParagraphAlignment enum.
    /// </summary>
    /// <param name="alignment">The alignment string (left, center, right).</param>
    /// <returns>The corresponding ParagraphAlignment enum value.</returns>
    private static ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }

    /// <summary>
    ///     Converts wrap type string to WrapType enum.
    /// </summary>
    /// <param name="wrapType">The wrap type string (inline, square, tight, through, topandbottom, none).</param>
    /// <returns>The corresponding WrapType enum value.</returns>
    private static WrapType GetWrapType(string wrapType)
    {
        return wrapType.ToLower() switch
        {
            "square" => WrapType.Square,
            "tight" => WrapType.Tight,
            "through" => WrapType.Through,
            "topandbottom" => WrapType.TopBottom,
            "none" => WrapType.None,
            _ => WrapType.Inline
        };
    }

    /// <summary>
    ///     Converts alignment string to HorizontalAlignment enum for floating images.
    /// </summary>
    /// <param name="alignment">The alignment string (left, center, right).</param>
    /// <returns>The corresponding HorizontalAlignment enum value.</returns>
    private static HorizontalAlignment GetHorizontalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => HorizontalAlignment.Left,
            "center" => HorizontalAlignment.Center,
            "right" => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Left
        };
    }

    /// <summary>
    ///     Converts alignment string to VerticalAlignment enum for floating images.
    /// </summary>
    /// <param name="alignment">The alignment string (top, center, bottom).</param>
    /// <returns>The corresponding VerticalAlignment enum value.</returns>
    private static VerticalAlignment GetVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => VerticalAlignment.Top,
            "center" => VerticalAlignment.Center,
            "bottom" => VerticalAlignment.Bottom,
            _ => VerticalAlignment.Top
        };
    }

    /// <summary>
    ///     Inserts a professional caption with automatic figure numbering.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="caption">The caption text.</param>
    /// <param name="alignment">The caption alignment.</param>
    private static void InsertCaption(DocumentBuilder builder, string caption, string alignment)
    {
        // Use professional Caption style with SEQ field for automatic figure numbering
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
        builder.ParagraphFormat.Alignment = GetAlignment(alignment);
        builder.Write("Figure ");
        builder.InsertField("SEQ Figure \\* ARABIC");
        builder.Write(": " + caption);
        builder.Writeln();
        // Reset to normal style after caption
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
    }
}