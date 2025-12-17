using System.Globalization;
using System.Text;
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
                description = "Image width in points (optional, for add/edit operations)"
            },
            height = new
            {
                type = "number",
                description = "Image height in points (optional, for add/edit operations)"
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
                description = "Alternative text for accessibility (optional, for edit operation)"
            },
            title = new
            {
                type = "string",
                description = "Image title (optional, for edit operation)"
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
            preservePosition = new
            {
                type = "boolean",
                description = "Preserve original image position and wrapping (default: true, for replace operation)"
            },
            prefix = new
            {
                type = "string",
                description = "Filename prefix for extracted images (optional, default: 'image', for extract operation)"
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

        SecurityHelper.ValidateFilePath(path);

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(arguments, path),
            "edit" => await EditImageAsync(arguments, path),
            "delete" => await DeleteImageAsync(arguments, path),
            "get" => await GetImagesAsync(arguments, path),
            "replace" => await ReplaceImageAsync(arguments, path),
            "extract" => await ExtractImagesAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing imagePath, optional width, height, position, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddImageAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
        var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
        var height = ArgumentHelper.GetDoubleNullable(arguments, "height");
        var alignment = ArgumentHelper.GetString(arguments, "alignment", "left");
        var textWrapping = ArgumentHelper.GetString(arguments, "textWrapping", "inline");
        var caption = ArgumentHelper.GetStringNullable(arguments, "caption");
        var captionPosition = ArgumentHelper.GetString(arguments, "captionPosition", "below");

        if (!File.Exists(imagePath)) throw new FileNotFoundException($"Image file not found: {imagePath}");

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        // Add caption above if specified
        if (!string.IsNullOrEmpty(caption) && captionPosition == "above")
        {
            builder.ParagraphFormat.Alignment = GetAlignment(alignment);
            builder.Font.Italic = true;
            builder.Writeln(caption);
            builder.Font.Italic = false;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        }

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

            // Set horizontal alignment for floating images
            if (alignment == "center")
            {
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.HorizontalAlignment = HorizontalAlignment.Center;
            }
            else if (alignment == "right")
            {
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.HorizontalAlignment = HorizontalAlignment.Right;
            }
            else
            {
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.HorizontalAlignment = HorizontalAlignment.Left;
            }
        }

        // Reset paragraph alignment only after caption (if any) is added

        // Add caption below if specified
        if (!string.IsNullOrEmpty(caption) && captionPosition == "below")
        {
            if (textWrapping == "inline")
                // For inline images, caption should be in a new paragraph with same alignment
                builder.ParagraphFormat.Alignment = GetAlignment(alignment);
            builder.Font.Italic = true;
            builder.Writeln(caption);
            builder.Font.Italic = false;
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
        if (!string.IsNullOrEmpty(caption)) result += $"Caption: {caption} ({captionPosition})\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Edits image properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing imageIndex, optional width, height, position, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> EditImageAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

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
                var hAlign = ArgumentHelper.GetStringNullable(arguments, "horizontalAlignment") ?? "left";
                if (!string.IsNullOrEmpty(hAlign))
                {
                    shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                    shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);
                }

                var vAlign = ArgumentHelper.GetStringNullable(arguments, "verticalAlignment") ?? "top";
                if (!string.IsNullOrEmpty(vAlign))
                {
                    shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                    shape.VerticalAlignment = GetVerticalAlignment(vAlign);
                }
            }
        }
        else if (shape.WrapType != WrapType.Inline)
        {
            var hAlign = ArgumentHelper.GetStringNullable(arguments, "horizontalAlignment") ?? "left";
            if (!string.IsNullOrEmpty(hAlign))
            {
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);
            }

            var vAlign = ArgumentHelper.GetStringNullable(arguments, "verticalAlignment") ?? "top";
            if (!string.IsNullOrEmpty(vAlign))
            {
                shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                shape.VerticalAlignment = GetVerticalAlignment(vAlign);
            }
        }

        // Apply alternative text
        var altText = ArgumentHelper.GetStringNullable(arguments, "alternativeText");
        if (!string.IsNullOrEmpty(altText))
            shape.AlternativeText = altText;

        // Apply title
        var title = ArgumentHelper.GetStringNullable(arguments, "title");
        if (!string.IsNullOrEmpty(title))
            shape.Title = title;

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

        var changesDesc = changes.Count > 0 ? string.Join(", ", changes) : "properties";

        return await Task.FromResult($"Image {imageIndex} edited successfully ({changesDesc}). Output: {outputPath}");
    }

    /// <summary>
    ///     Deletes an image from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing imageIndex, optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteImageAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

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
            catch
            {
                // Size information may not be available, but this is not critical
                // Continue without the size information
            }

        shapeToDelete.Remove();

        doc.Save(outputPath);

        var remainingCount = GetAllImages(doc, sectionIndex).Count;

        var result = $"{imageInfo} deleted successfully\n";
        result += $"Remaining images in document: {remainingCount}\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Gets all images from the document
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all images</returns>
    private async Task<string> GetImagesAsync(JsonObject? arguments, string path)
    {
        var doc = new Document(path);
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", -1);

        var shapes = GetAllImages(doc, sectionIndex);

        var result = new StringBuilder();

        result.AppendLine("=== Document Image Information ===\n");
        if (sectionIndex == -1)
            result.AppendLine($"Total images: {shapes.Count}\n");
        else
            result.AppendLine($"Section {sectionIndex} images: {shapes.Count}\n");

        if (shapes.Count == 0)
        {
            result.AppendLine("No images found");
            if (sectionIndex != -1)
                result.AppendLine(
                    $"(No images found in section {sectionIndex}, use sectionIndex=-1 to search all sections)");
            return await Task.FromResult(result.ToString());
        }

        for (var i = 0; i < shapes.Count; i++)
        {
            var shape = shapes[i];
            result.AppendLine($"[Image {i}]");
            result.AppendLine($"Name: {shape.Name ?? "(No name)"}");
            result.AppendLine($"Width: {shape.Width} points");
            result.AppendLine($"Height: {shape.Height} points");

            if (shape.IsInline)
            {
                // For inline images, show paragraph alignment instead of position
                if (shape.ParentNode is Paragraph parentPara)
                {
                    result.AppendLine($"Alignment: {parentPara.ParagraphFormat.Alignment} (paragraph alignment)");
                    result.AppendLine("Position: Inline in paragraph (X/Y position not applicable for inline images)");
                }
                else
                {
                    result.AppendLine($"Position: X={shape.Left}, Y={shape.Top}");
                }
            }
            else
            {
                // For floating images, show position and alignment
                result.AppendLine($"Position: X={shape.Left}, Y={shape.Top}");
                result.AppendLine($"Horizontal alignment: {shape.HorizontalAlignment}");
                result.AppendLine($"Vertical alignment: {shape.VerticalAlignment}");
                result.AppendLine($"Text wrapping: {shape.WrapType}");
            }

            if (shape.ImageData != null)
            {
                result.AppendLine($"Image type: {shape.ImageData.ImageType}");
                var imageSize = shape.ImageData.ImageSize;
                result.AppendLine($"Original size: {imageSize.WidthPixels} × {imageSize.HeightPixels} pixels");
            }

            result.AppendLine($"Is inline: {shape.IsInline}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    ///     Replaces an existing image with a new one
    /// </summary>
    /// <param name="arguments">JSON arguments containing imageIndex, newImagePath, optional outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message</returns>
    private async Task<string> ReplaceImageAsync(JsonObject? arguments, string path)
    {
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
        var newImagePath = ArgumentHelper.GetString(arguments, "newImagePath");
        var preserveSize = ArgumentHelper.GetBool(arguments, "preserveSize");
        var preservePosition = ArgumentHelper.GetBool(arguments, "preservePosition");
        var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        SecurityHelper.ValidateFilePath(newImagePath, "newImagePath");

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
                shapeToReplace.Width = originalWidth;
                shapeToReplace.Height = originalHeight;
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
        if (preserveSize) result += $"Preserved size: {originalWidth:F1} pt x {originalHeight:F1} pt\n";
        if (preservePosition) result += "Preserved position and wrapping\n";
        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Extracts images from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing outputDirectory, optional imageIndex</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Success message with extracted image count</returns>
    private async Task<string> ExtractImagesAsync(JsonObject? arguments, string path)
    {
        var outputDir = ArgumentHelper.GetString(arguments, "outputDir");
        var prefix = ArgumentHelper.GetString(arguments, "prefix", "image");

        SecurityHelper.ValidateFilePath(outputDir, "outputDir");

        Directory.CreateDirectory(outputDir);

        var doc = new Document(path);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();

        if (shapes.Count == 0) return await Task.FromResult("No images found in document");

        var extractedFiles = new List<string>();

        for (var i = 0; i < shapes.Count; i++)
        {
            var shape = shapes[i];
            var imageData = shape.ImageData;

            var imageBytes = imageData.ImageBytes;
            var extension = "img";

            if (imageBytes is { Length: > 4 })
            {
                if (imageBytes[0] == 0xFF && imageBytes[1] == 0xD8)
                    extension = "jpg";
                else if (imageBytes[0] == 0x89 && imageBytes[1] == 0x50 && imageBytes[2] == 0x4E &&
                         imageBytes[3] == 0x47)
                    extension = "png";
                else if (imageBytes[0] == 0x42 && imageBytes[1] == 0x4D)
                    extension = "bmp";
                else if (imageBytes[0] == 0x47 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46)
                    extension = "gif";
            }

            var safePrefix = SecurityHelper.SanitizeFileName(prefix);
            var filename = $"{safePrefix}_{i + 1:D3}.{extension}";
            var outputPath = Path.Combine(outputDir, filename);

            await using (var stream = File.Create(outputPath))
            {
                imageData.Save(stream);
            }

            extractedFiles.Add(outputPath);
        }

        return await Task.FromResult($"Successfully extracted {shapes.Count} images to: {outputDir}\n" +
                                     $"File list:\n" + string.Join("\n",
                                         extractedFiles.Select(f => $"  - {Path.GetFileName(f)}")));
    }

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

    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };
    }

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
}