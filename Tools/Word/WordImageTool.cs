using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Word images (add, edit, delete, get, replace, extract)
/// Merges: WordAddImageTool, WordEditImageTool, WordDeleteImageTool, WordGetImagesTool, WordReplaceImageTool, WordExtractImagesTool
/// </summary>
public class WordImageTool : IAsposeTool
{
    public string Description => "Manage Word document images: add, edit, delete, get all, replace, or extract";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'edit', 'delete', 'get', 'replace', 'extract'",
                @enum = new[] { "add", "edit", "delete", "get", "replace", "extract" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for add/edit/delete/replace operations)"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory (required for extract operation)"
            },
            imagePath = new
            {
                type = "string",
                description = "Image file path (required for add/replace operations)"
            },
            imageIndex = new
            {
                type = "number",
                description = "Image index (0-based, required for edit/delete/replace operations)"
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
                description = "Text wrapping: inline, square, tight, through, topAndBottom, none (optional, for add/edit operations)",
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
                description = "Horizontal alignment for floating images: left, center, right (optional, for edit operation)",
                @enum = new[] { "left", "center", "right" }
            },
            verticalAlignment = new
            {
                type = "string",
                description = "Vertical alignment for floating images: top, center, bottom (optional, for edit operation)",
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

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

    private async Task<string> AddImageAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required for add operation");
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();
        var alignment = arguments?["alignment"]?.GetValue<string>() ?? "left";
        var textWrapping = arguments?["textWrapping"]?.GetValue<string>() ?? "inline";
        var caption = arguments?["caption"]?.GetValue<string>();
        var captionPosition = arguments?["captionPosition"]?.GetValue<string>() ?? "below";

        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"找不到圖片檔案: {imagePath}");
        }

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

        // Set paragraph alignment for image
        builder.ParagraphFormat.Alignment = GetAlignment(alignment);

        // Insert image
        Shape shape;
        if (textWrapping == "inline")
        {
            shape = builder.InsertImage(imagePath);
        }
        else
        {
            shape = builder.InsertImage(imagePath);
            shape.WrapType = GetWrapType(textWrapping);
            
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
        }

        // Set size if specified
        if (width.HasValue)
            shape.Width = width.Value;
        
        if (height.HasValue)
            shape.Height = height.Value;

        // Reset paragraph alignment
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Add caption below if specified
        if (!string.IsNullOrEmpty(caption) && captionPosition == "below")
        {
            builder.ParagraphFormat.Alignment = GetAlignment(alignment);
            builder.Font.Italic = true;
            builder.Writeln(caption);
            builder.Font.Italic = false;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        }

        doc.Save(outputPath);

        var result = $"成功添加圖片\n";
        result += $"圖片: {Path.GetFileName(imagePath)}\n";
        if (width.HasValue || height.HasValue)
        {
            result += $"尺寸: {(width.HasValue ? width.Value.ToString() : "auto")} x {(height.HasValue ? height.Value.ToString() : "auto")} pt\n";
        }
        result += $"對齊: {alignment}\n";
        result += $"文繞圖: {textWrapping}\n";
        if (!string.IsNullOrEmpty(caption))
        {
            result += $"圖片說明: {caption} ({captionPosition})\n";
        }
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> EditImageAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required for edit operation");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        
        var allImages = GetAllImages(doc, sectionIndex);
        
        if (imageIndex < 0 || imageIndex >= allImages.Count)
        {
            throw new ArgumentException($"圖片索引 {imageIndex} 超出範圍 (文檔共有 {allImages.Count} 張圖片)");
        }
        
        var shape = allImages[imageIndex];
        
        // Apply size properties
        if (arguments?["width"] != null)
        {
            var width = arguments["width"]?.GetValue<double>();
            if (width.HasValue)
                shape.Width = width.Value;
        }
        
        if (arguments?["height"] != null)
        {
            var height = arguments["height"]?.GetValue<double>();
            if (height.HasValue)
                shape.Height = height.Value;
        }
        
        if (arguments?["aspectRatioLocked"] != null)
        {
            shape.AspectRatioLocked = arguments["aspectRatioLocked"]?.GetValue<bool>() ?? true;
        }
        
        // Apply alignment (for inline images)
        if (arguments?["alignment"] != null)
        {
            var alignment = arguments["alignment"]?.GetValue<string>() ?? "left";
            var parentPara = shape.ParentNode as Paragraph;
            if (parentPara != null)
            {
                parentPara.ParagraphFormat.Alignment = GetAlignment(alignment);
            }
        }
        
        // Apply text wrapping
        if (arguments?["textWrapping"] != null)
        {
            var textWrapping = arguments["textWrapping"]?.GetValue<string>() ?? "inline";
            shape.WrapType = GetWrapType(textWrapping);
            
            if (textWrapping != "inline")
            {
                if (arguments?["horizontalAlignment"] != null)
                {
                    var hAlign = arguments["horizontalAlignment"]?.GetValue<string>() ?? "left";
                    shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                    shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);
                }
                
                if (arguments?["verticalAlignment"] != null)
                {
                    var vAlign = arguments["verticalAlignment"]?.GetValue<string>() ?? "top";
                    shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                    shape.VerticalAlignment = GetVerticalAlignment(vAlign);
                }
            }
        }
        else if (shape.WrapType != WrapType.Inline)
        {
            if (arguments?["horizontalAlignment"] != null)
            {
                var hAlign = arguments["horizontalAlignment"]?.GetValue<string>() ?? "left";
                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.HorizontalAlignment = GetHorizontalAlignment(hAlign);
            }
            
            if (arguments?["verticalAlignment"] != null)
            {
                var vAlign = arguments["verticalAlignment"]?.GetValue<string>() ?? "top";
                shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                shape.VerticalAlignment = GetVerticalAlignment(vAlign);
            }
        }
        
        // Apply alternative text
        if (arguments?["alternativeText"] != null)
        {
            var altText = arguments["alternativeText"]?.GetValue<string>();
            shape.AlternativeText = altText ?? "";
        }
        
        // Apply title
        if (arguments?["title"] != null)
        {
            var title = arguments["title"]?.GetValue<string>();
            shape.Title = title ?? "";
        }
        
        doc.Save(outputPath);
        
        var changes = new List<string>();
        if (arguments?["width"] != null) changes.Add($"寬度：{arguments["width"]?.GetValue<double>()}");
        if (arguments?["height"] != null) changes.Add($"高度：{arguments["height"]?.GetValue<double>()}");
        if (arguments?["alignment"] != null) changes.Add($"對齊：{arguments["alignment"]?.GetValue<string>()}");
        if (arguments?["textWrapping"] != null) changes.Add($"文字環繞：{arguments["textWrapping"]?.GetValue<string>()}");
        
        var changesDesc = changes.Count > 0 ? string.Join("、", changes) : "屬性";
        
        return await Task.FromResult($"成功編輯圖片 {imageIndex} 的{changesDesc}。輸出: {outputPath}");
    }

    private async Task<string> DeleteImageAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required for delete operation");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        var doc = new Document(path);
        
        var allImages = GetAllImages(doc, sectionIndex);
        
        if (imageIndex < 0 || imageIndex >= allImages.Count)
        {
            throw new ArgumentException($"圖片索引 {imageIndex} 超出範圍 (文檔共有 {allImages.Count} 張圖片)");
        }
        
        var shapeToDelete = allImages[imageIndex];
        
        string imageInfo = $"圖片 #{imageIndex}";
        if (shapeToDelete.HasImage)
        {
            try
            {
                imageInfo += $" (寬度: {shapeToDelete.Width:F1} pt, 高度: {shapeToDelete.Height:F1} pt)";
            }
            catch { }
        }
        
        shapeToDelete.Remove();
        
        doc.Save(outputPath);
        
        int remainingCount = GetAllImages(doc, sectionIndex).Count;

        var result = $"成功刪除 {imageInfo}\n";
        result += $"文檔剩餘圖片數: {remainingCount}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> GetImagesAsync(JsonObject? arguments, string path)
    {
        var doc = new Document(path);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.HasImage)
            .ToList();
        
        var result = new StringBuilder();

        result.AppendLine("=== 文檔圖片資訊 ===\n");
        result.AppendLine($"總圖片數: {shapes.Count}\n");

        if (shapes.Count == 0)
        {
            result.AppendLine("未找到圖片");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < shapes.Count; i++)
        {
            var shape = shapes[i];
            result.AppendLine($"【圖片 {i}】");
            result.AppendLine($"名稱: {shape.Name ?? "(無名稱)"}");
            result.AppendLine($"寬度: {shape.Width} 點");
            result.AppendLine($"高度: {shape.Height} 點");
            result.AppendLine($"位置: X={shape.Left}, Y={shape.Top}");
            
            if (shape.ImageData != null)
            {
                result.AppendLine($"圖片類型: {shape.ImageData.ImageType}");
                var imageSize = shape.ImageData.ImageSize;
                result.AppendLine($"原始尺寸: {imageSize.WidthPixels} × {imageSize.HeightPixels} 像素");
            }
            
            result.AppendLine($"是否在文本內: {shape.IsInline}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> ReplaceImageAsync(JsonObject? arguments, string path)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required for replace operation");
        var newImagePath = arguments?["newImagePath"]?.GetValue<string>() ?? throw new ArgumentException("newImagePath is required for replace operation");
        var preserveSize = arguments?["preserveSize"]?.GetValue<bool>() ?? true;
        var preservePosition = arguments?["preservePosition"]?.GetValue<bool>() ?? true;
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        SecurityHelper.ValidateFilePath(newImagePath, "newImagePath");

        if (!File.Exists(newImagePath))
        {
            throw new FileNotFoundException($"找不到圖片檔案: {newImagePath}");
        }

        var doc = new Document(path);
        
        var allImages = GetAllImages(doc, sectionIndex);
        
        if (imageIndex < 0 || imageIndex >= allImages.Count)
        {
            throw new ArgumentException($"圖片索引 {imageIndex} 超出範圍 (文檔共有 {allImages.Count} 張圖片)");
        }
        
        var shapeToReplace = allImages[imageIndex];
        
        double originalWidth = shapeToReplace.Width;
        double originalHeight = shapeToReplace.Height;
        WrapType originalWrapType = shapeToReplace.WrapType;
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
            throw new InvalidOperationException($"替換圖片時發生錯誤: {ex.Message}", ex);
        }
        
        doc.Save(outputPath);
        
        var result = $"成功替換圖片 #{imageIndex}\n";
        result += $"新圖片: {Path.GetFileName(newImagePath)}\n";
        if (preserveSize)
        {
            result += $"保留尺寸: {originalWidth:F1} pt x {originalHeight:F1} pt\n";
        }
        if (preservePosition)
        {
            result += $"保留位置和環繞方式\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> ExtractImagesAsync(JsonObject? arguments, string path)
    {
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? throw new ArgumentException("outputDir is required for extract operation");
        var prefix = arguments?["prefix"]?.GetValue<string>() ?? "image";

        SecurityHelper.ValidateFilePath(outputDir, "outputDir");

        Directory.CreateDirectory(outputDir);

        var doc = new Document(path);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
        
        if (shapes.Count == 0)
        {
            return await Task.FromResult("文檔中沒有找到圖片");
        }

        var extractedFiles = new List<string>();
        
        for (int i = 0; i < shapes.Count; i++)
        {
            var shape = shapes[i];
            var imageData = shape.ImageData;
            
            var imageBytes = imageData.ImageBytes;
            string extension = "img";
            
            if (imageBytes != null && imageBytes.Length > 4)
            {
                if (imageBytes[0] == 0xFF && imageBytes[1] == 0xD8)
                    extension = "jpg";
                else if (imageBytes[0] == 0x89 && imageBytes[1] == 0x50 && imageBytes[2] == 0x4E && imageBytes[3] == 0x47)
                    extension = "png";
                else if (imageBytes[0] == 0x42 && imageBytes[1] == 0x4D)
                    extension = "bmp";
                else if (imageBytes[0] == 0x47 && imageBytes[1] == 0x49 && imageBytes[2] == 0x46)
                    extension = "gif";
            }

            var safePrefix = SecurityHelper.SanitizeFileName(prefix);
            var filename = $"{safePrefix}_{i + 1:D3}.{extension}";
            var outputPath = Path.Combine(outputDir, filename);
            
            using (var stream = File.Create(outputPath))
            {
                imageData.Save(stream);
            }
            
            extractedFiles.Add(outputPath);
        }

        return await Task.FromResult($"成功提取 {shapes.Count} 張圖片到: {outputDir}\n" +
                                    $"檔案列表:\n" + string.Join("\n", extractedFiles.Select(f => $"  - {Path.GetFileName(f)}")));
    }

    private List<Shape> GetAllImages(Document doc, int sectionIndex)
    {
        List<Shape> allImages = new List<Shape>();
        
        if (sectionIndex == -1)
        {
            foreach (Section section in doc.Sections)
            {
                var shapes = section.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).ToList();
                allImages.AddRange(shapes);
            }
        }
        else
        {
            if (sectionIndex >= doc.Sections.Count)
            {
                throw new ArgumentException($"節索引 {sectionIndex} 超出範圍 (文檔共有 {doc.Sections.Count} 個節)");
            }
            
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

