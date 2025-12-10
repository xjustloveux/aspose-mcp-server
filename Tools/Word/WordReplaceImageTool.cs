using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordReplaceImageTool : IAsposeTool
{
    public string Description => "Replace an existing image in Word document while preserving position and size";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            imageIndex = new
            {
                type = "number",
                description = "Image index (0-based, from word_get_content_detailed or word_extract_images)"
            },
            newImagePath = new
            {
                type = "string",
                description = "Path to the new image file"
            },
            preserveSize = new
            {
                type = "boolean",
                description = "Preserve original image size (default: true)"
            },
            preservePosition = new
            {
                type = "boolean",
                description = "Preserve original image position and wrapping (default: true)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0. Use -1 to search all sections"
            }
        },
        required = new[] { "path", "imageIndex", "newImagePath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required");
        var newImagePath = arguments?["newImagePath"]?.GetValue<string>() ?? throw new ArgumentException("newImagePath is required");
        var preserveSize = arguments?["preserveSize"]?.GetValue<bool>() ?? true;
        var preservePosition = arguments?["preservePosition"]?.GetValue<bool>() ?? true;
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        if (!File.Exists(newImagePath))
        {
            throw new FileNotFoundException($"找不到圖片文件: {newImagePath}");
        }

        var doc = new Document(path);
        
        // Get all images (shapes with images) from document
        List<Shape> allImages = new List<Shape>();
        
        if (sectionIndex == -1)
        {
            // Search all sections
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
        
        if (imageIndex < 0 || imageIndex >= allImages.Count)
        {
            throw new ArgumentException($"圖片索引 {imageIndex} 超出範圍 (文檔共有 {allImages.Count} 張圖片)");
        }
        
        var shapeToReplace = allImages[imageIndex];
        
        // Save original properties
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
        
        // Replace image data
        try
        {
            shapeToReplace.ImageData.SetImage(newImagePath);
            
            // Restore size if requested
            if (preserveSize)
            {
                shapeToReplace.Width = originalWidth;
                shapeToReplace.Height = originalHeight;
            }
            
            // Restore position and wrapping if requested
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
}

