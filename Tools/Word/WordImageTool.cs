using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing Word images (add, edit, delete, get, replace, extract)
///     Merges: WordAddImageTool, WordEditImageTool, WordDeleteImageTool, WordGetImagesTool, WordReplaceImageTool,
///     WordExtractImagesTool
/// </summary>
public class WordImageTool : IAsposeTool
{
    public string Description =>
        @"Manage Word document images. Supports 6 operations: add, edit, delete, get, replace, extract.

Usage examples:
- Add image: word_image(operation='add', path='doc.docx', imagePath='image.png', width=200)
- Edit image: word_image(operation='edit', path='doc.docx', imageIndex=0, width=300, height=200)
- Delete image: word_image(operation='delete', path='doc.docx', imageIndex=0)
- Get all images: word_image(operation='get', path='doc.docx')
- Replace image: word_image(operation='replace', path='doc.docx', imageIndex=0, imagePath='new_image.png')
- Extract images: word_image(operation='extract', path='doc.docx', outputDir='images/')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a new image (required params: path, imagePath)
- 'edit': Edit existing image (required params: path, imageIndex)
- 'delete': Delete an image (required params: path, imageIndex)
- 'get': Get all images info (required params: path)
- 'replace': Replace an image (required params: path, imageIndex, imagePath)
- 'extract': Extract all images (required params: path, outputDir)",
                @enum = new[] { "add", "edit", "delete", "get", "replace", "extract" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (if not provided, overwrites input, for add/edit/delete/replace operations)"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory (required for extract operation)"
            },
            imagePath = new
            {
                type = "string",
                description = "Image file path (required for add and replace operations)"
            },
            imageIndex = new
            {
                type = "number",
                description =
                    "Image index (0-based, required for edit, delete, and replace operations). Note: After delete operations, subsequent image indices will shift automatically. Use 'get' operation to refresh indices."
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0, use -1 to search all sections)"
            },
            width = new
            {
                type = "number",
                description = "Image width in points (72 pts = 1 inch, optional, for add/edit operations)"
            },
            height = new
            {
                type = "number",
                description = "Image height in points (72 pts = 1 inch, optional, for add/edit operations)"
            },
            alignment = new
            {
                type = "string",
                description = "Horizontal alignment: left, center, right (optional, for add/edit operations)",
                @enum = new[] { "left", "center", "right" }
            },
            textWrapping = new
            {
                type = "string",
                description =
                    "Text wrapping: inline, square, tight, through, topAndBottom, none (optional, for add/edit operations)",
                @enum = new[] { "inline", "square", "tight", "through", "topAndBottom", "none" }
            },
            caption = new
            {
                type = "string",
                description = "Image caption text (optional, for add operation)"
            },
            captionPosition = new
            {
                type = "string",
                description = "Caption position: above, below (optional, for add operation)",
                @enum = new[] { "above", "below" }
            },
            aspectRatioLocked = new
            {
                type = "boolean",
                description = "Lock aspect ratio (optional, for edit operation)"
            },
            horizontalAlignment = new
            {
                type = "string",
                description =
                    "Horizontal alignment for floating images: left, center, right (optional, for edit operation)",
                @enum = new[] { "left", "center", "right" }
            },
            verticalAlignment = new
            {
                type = "string",
                description =
                    "Vertical alignment for floating images: top, center, bottom (optional, for edit operation)",
                @enum = new[] { "top", "center", "bottom" }
            },
            alternativeText = new
            {
                type = "string",
                description = "Alternative text for accessibility (optional, for add/edit operation)"
            },
            title = new
            {
                type = "string",
                description = "Image title (optional, for add/edit operation)"
            },
            linkUrl = new
            {
                type = "string",
                description =
                    "Hyperlink URL for the image. When clicked, opens the specified URL (optional, for add/edit operation). Use empty string to remove existing hyperlink."
            },
            newImagePath = new
            {
                type = "string",
                description = "New image file path (required for replace operation)"
            },
            preserveSize = new
            {
                type = "boolean",
                description = "Preserve original image size (default: true, for replace operation)"
            },
            smartFit = new
            {
                type = "boolean",
                description =
                    "When true, keeps original width and calculates height proportionally based on new image aspect ratio (avoids distortion when aspect ratios differ, default: false, for replace operation). Only applies when preserveSize is true."
            },
            preservePosition = new
            {
                type = "boolean",
                description = "Preserve original image position and wrapping (default: true, for replace operation)"
            },
            prefix = new
            {
                type = "string",
                description = "Filename prefix for extracted images (optional, default: 'image', for extract operation)"
            },
            extractImageIndex = new
            {
                type = "number",
                description =
                    "Specific image index to extract (0-based, optional, for extract operation). If not provided, extracts all images."
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(path, outputPath, arguments),
            "edit" => await EditImageAsync(path, outputPath, arguments),
            "delete" => await DeleteImageAsync(path, outputPath, arguments),
            "get" => await GetImagesAsync(path, arguments),
            "replace" => await ReplaceImageAsync(path, outputPath, arguments),
            "extract" => await ExtractImagesAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing imagePath, optional width, height, alignment, textWrapping</param>
    /// <returns>Success message</returns>
    private Task<string> AddImageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
            var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
            var height = ArgumentHelper.GetDoubleNullable(arguments, "height");
            var alignment = ArgumentHelper.GetString(arguments, "alignment", "left");
            var textWrapping = ArgumentHelper.GetString(arguments, "textWrapping", "inline");
            var caption = ArgumentHelper.GetStringNullable(arguments, "caption");
            var captionPosition = ArgumentHelper.GetString(arguments, "captionPosition", "below");
            var linkUrl = ArgumentHelper.GetStringNullable(arguments, "linkUrl");
            var alternativeText = ArgumentHelper.GetStringNullable(arguments, "alternativeText");
            var title = ArgumentHelper.GetStringNullable(arguments, "title");

            if (!File.Exists(imagePath)) throw new FileNotFoundException($"Image file not found: {imagePath}");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            // Add caption above if specified (using professional Caption style with SEQ field)
            if (!string.IsNullOrEmpty(caption) && captionPosition == "above")
                InsertCaption(builder, caption, alignment);

            // Insert image
            Shape shape;
            if (textWrapping == "inline")
            {
                // For inline images, alignment is controlled by paragraph alignment
                // Set paragraph alignment before inserting image
                var paraAlignment = GetAlignment(alignment);
                builder.ParagraphFormat.Alignment = paraAlignment;
                shape = builder.InsertImage(imagePath);

                // Set size if specified
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

                // Keep paragraph alignment for inline images
                builder.ParagraphFormat.Alignment = paraAlignment;
            }
            else
            {
                // For floating images, use shape positioning
                shape = builder.InsertImage(imagePath);
                shape.WrapType = GetWrapType(textWrapping);

                // Set size if specified
                if (width.HasValue)
                    shape.Width = width.Value;

                if (height.HasValue)
                    shape.Height = height.Value;

                // Set alignment for floating images (relative to Column/Paragraph for better text flow)
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
                shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
                if (alignment == "center")
                    shape.HorizontalAlignment = HorizontalAlignment.Center;
                else if (alignment == "right")
                    shape.HorizontalAlignment = HorizontalAlignment.Right;
                else
                    shape.HorizontalAlignment = HorizontalAlignment.Left;
            }

            // Set hyperlink if provided
            if (!string.IsNullOrEmpty(linkUrl))
                shape.HRef = linkUrl;

            // Set alternative text if provided
            if (!string.IsNullOrEmpty(alternativeText))
                shape.AlternativeText = alternativeText;

            // Set title if provided
            if (!string.IsNullOrEmpty(title))
                shape.Title = title;

            // Reset paragraph alignment only after caption (if any) is added

            // Add caption below if specified (using professional Caption style with SEQ field)
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

            doc.Save(outputPath);

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
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Edits image properties
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing imageIndex, optional width, height, alignment, textWrapping</param>
    /// <returns>Success message</returns>
    private Task<string> EditImageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);

            var allImages = GetAllImages(doc, sectionIndex);

            if (imageIndex < 0 || imageIndex >= allImages.Count)
                throw new ArgumentException(
                    $"Image index {imageIndex} is out of range (document has {allImages.Count} images)");

            var shape = allImages[imageIndex];

            // Apply size properties
            var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
            if (width.HasValue)
                shape.Width = width.Value;

            var height = ArgumentHelper.GetDoubleNullable(arguments, "height");
            if (height.HasValue)
                shape.Height = height.Value;

            var aspectRatioLocked = ArgumentHelper.GetBoolNullable(arguments, "aspectRatioLocked");
            if (aspectRatioLocked.HasValue)
                shape.AspectRatioLocked = aspectRatioLocked.Value;

            // Apply alignment (for inline images)
            var alignment = ArgumentHelper.GetStringNullable(arguments, "alignment") ?? "left";
            if (!string.IsNullOrEmpty(alignment))
                if (shape.ParentNode is Paragraph parentPara)
                    parentPara.ParagraphFormat.Alignment = GetAlignment(alignment);

            // Apply text wrapping
            var textWrapping = ArgumentHelper.GetStringNullable(arguments, "textWrapping") ?? "inline";
            if (!string.IsNullOrEmpty(textWrapping))
            {
                shape.WrapType = GetWrapType(textWrapping);

                if (textWrapping != "inline")
                {
                    // Use Column/Paragraph positioning for better text flow
                    shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
                    shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

                    var hAlign = ArgumentHelper.GetStringNullable(arguments, "horizontalAlignment") ?? "left";
                    if (!string.IsNullOrEmpty(hAlign)) shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);

                    var vAlign = ArgumentHelper.GetStringNullable(arguments, "verticalAlignment") ?? "top";
                    if (!string.IsNullOrEmpty(vAlign)) shape.VerticalAlignment = GetVerticalAlignment(vAlign);
                }
            }
            else if (shape.WrapType != WrapType.Inline)
            {
                // Use Column/Paragraph positioning for better text flow
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
                shape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

                var hAlign = ArgumentHelper.GetStringNullable(arguments, "horizontalAlignment") ?? "left";
                if (!string.IsNullOrEmpty(hAlign)) shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);

                var vAlign = ArgumentHelper.GetStringNullable(arguments, "verticalAlignment") ?? "top";
                if (!string.IsNullOrEmpty(vAlign)) shape.VerticalAlignment = GetVerticalAlignment(vAlign);
            }

            // Apply alternative text
            var altText = ArgumentHelper.GetStringNullable(arguments, "alternativeText");
            if (!string.IsNullOrEmpty(altText))
                shape.AlternativeText = altText;

            // Apply title
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            if (!string.IsNullOrEmpty(title))
                shape.Title = title;

            // Apply hyperlink
            var linkUrl = ArgumentHelper.GetStringNullable(arguments, "linkUrl");
            if (linkUrl != null)
                // Note: HRef property doesn't accept null, use empty string to clear
                shape.HRef = linkUrl;

            doc.Save(outputPath);

            var changes = new List<string>();
            var widthValue = ArgumentHelper.GetDoubleNullable(arguments, "width");
            var heightValue = ArgumentHelper.GetDoubleNullable(arguments, "height");
            var alignmentValue = ArgumentHelper.GetStringNullable(arguments, "alignment");
            var textWrappingValue = ArgumentHelper.GetStringNullable(arguments, "textWrapping");
            if (widthValue.HasValue) changes.Add($"Width: {widthValue.Value}");
            if (heightValue.HasValue) changes.Add($"Height: {heightValue.Value}");
            if (alignmentValue != null) changes.Add($"Alignment: {alignmentValue}");
            if (textWrappingValue != null) changes.Add($"Text wrapping: {textWrappingValue}");
            if (linkUrl != null)
                changes.Add(string.IsNullOrEmpty(linkUrl) ? "Hyperlink: removed" : $"Hyperlink: {linkUrl}");
            if (altText != null) changes.Add($"Alt text: {altText}");
            if (title != null) changes.Add($"Title: {title}");

            var changesDesc = changes.Count > 0 ? string.Join(", ", changes) : "properties";

            return $"Image {imageIndex} edited successfully ({changesDesc}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes an image from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing imageIndex, optional sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteImageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);

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

            doc.Save(outputPath);

            var remainingCount = GetAllImages(doc, sectionIndex).Count;

            var result = $"{imageInfo} deleted successfully\n";
            result += $"Remaining images in document: {remainingCount}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Gets all images from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="arguments">JSON arguments containing optional sectionIndex</param>
    /// <returns>JSON formatted string with all images</returns>
    private Task<string> GetImagesAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", -1);

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

            var imageList = new List<object>();
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
        });
    }

    /// <summary>
    ///     Replaces an existing image with a new one
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing imageIndex, newImagePath, optional preserveSize, smartFit</param>
    /// <returns>Success message</returns>
    private Task<string> ReplaceImageAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
            var newImagePath = ArgumentHelper.GetString(arguments, "newImagePath");
            var preserveSize = ArgumentHelper.GetBool(arguments, "preserveSize", true);
            var smartFit = ArgumentHelper.GetBool(arguments, "smartFit", false);
            var preservePosition = ArgumentHelper.GetBool(arguments, "preservePosition", true);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            SecurityHelper.ValidateFilePath(newImagePath, "newImagePath", true);

            if (!File.Exists(newImagePath)) throw new FileNotFoundException($"Image file not found: {newImagePath}");

            var doc = new Document(path);

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

            doc.Save(outputPath);

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
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Extracts images from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="arguments">JSON arguments containing outputDir, optional prefix, extractImageIndex</param>
    /// <returns>Success message with extracted image count</returns>
    private Task<string> ExtractImagesAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var outputDir = ArgumentHelper.GetString(arguments, "outputDir");
            var prefix = ArgumentHelper.GetString(arguments, "prefix", "image");
            var extractImageIndex = ArgumentHelper.GetIntNullable(arguments, "extractImageIndex");

            SecurityHelper.ValidateFilePath(outputDir, "outputDir", true);

            Directory.CreateDirectory(outputDir);

            var doc = new Document(path);
            var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();

            if (shapes.Count == 0) return "No images found in document";

            // Validate extractImageIndex if provided
            if (extractImageIndex.HasValue)
                if (extractImageIndex.Value < 0 || extractImageIndex.Value >= shapes.Count)
                    throw new ArgumentException(
                        $"Image index {extractImageIndex.Value} is out of range (document has {shapes.Count} images)");

            var extractedFiles = new List<string>();

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
                var outputPath = Path.Combine(outputDir, filename);

                using (var stream = File.Create(outputPath))
                {
                    imageData.Save(stream);
                }

                extractedFiles.Add(outputPath);
            }

            if (extractImageIndex.HasValue)
                return $"Successfully extracted image #{extractImageIndex.Value} to: {outputDir}\n" +
                       $"File: {Path.GetFileName(extractedFiles[0])}";

            return $"Successfully extracted {shapes.Count} images to: {outputDir}\n" +
                   $"File list:\n" + string.Join("\n",
                       extractedFiles.Select(f => $"  - {Path.GetFileName(f)}"));
        });
    }

    /// <summary>
    ///     Gets all images from the document or a specific section
    /// </summary>
    /// <param name="doc">Word document</param>
    /// <param name="sectionIndex">Section index (-1 for all sections)</param>
    /// <returns>List of Shape objects containing images</returns>
    private List<Shape> GetAllImages(Document doc, int sectionIndex)
    {
        var allImages = new List<Shape>();

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
    ///     Converts alignment string to ParagraphAlignment enum
    /// </summary>
    /// <param name="alignment">Alignment string (left, center, right)</param>
    /// <returns>ParagraphAlignment enum value</returns>
    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }

    /// <summary>
    ///     Converts wrap type string to WrapType enum
    /// </summary>
    /// <param name="wrapType">Wrap type string (inline, square, tight, through, topAndBottom, none)</param>
    /// <returns>WrapType enum value</returns>
    private WrapType GetWrapType(string wrapType)
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
    ///     Converts alignment string to HorizontalAlignment enum for floating images
    /// </summary>
    /// <param name="alignment">Alignment string (left, center, right)</param>
    /// <returns>HorizontalAlignment enum value</returns>
    private HorizontalAlignment GetHorizontalAlignment(string alignment)
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
    ///     Converts alignment string to VerticalAlignment enum for floating images
    /// </summary>
    /// <param name="alignment">Alignment string (top, center, bottom)</param>
    /// <returns>VerticalAlignment enum value</returns>
    private VerticalAlignment GetVerticalAlignment(string alignment)
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
    ///     Inserts a professional caption with automatic figure numbering
    /// </summary>
    /// <param name="builder">DocumentBuilder for inserting content</param>
    /// <param name="caption">Caption text</param>
    /// <param name="alignment">Caption alignment (left, center, right)</param>
    private void InsertCaption(DocumentBuilder builder, string caption, string alignment)
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