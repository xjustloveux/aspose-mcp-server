using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordDeleteImageTool : IAsposeTool
{
    public string Description => "Delete a specific image from Word document by index";

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
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0. Use -1 to search all sections"
            }
        },
        required = new[] { "path", "imageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

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
        
        var shapeToDelete = allImages[imageIndex];
        
        // Get image info before deletion (for confirmation message)
        string imageInfo = $"圖片 #{imageIndex}";
        if (shapeToDelete.HasImage)
        {
            try
            {
                var imageData = shapeToDelete.ImageData;
                if (imageData != null)
                {
                    imageInfo += $" (寬度: {shapeToDelete.Width:F1} pt, 高度: {shapeToDelete.Height:F1} pt)";
                }
            }
            catch
            {
                // Ignore errors when getting image info
            }
        }
        
        // Delete the image (shape)
        shapeToDelete.Remove();
        
        doc.Save(outputPath);
        
        // Count remaining images
        int remainingCount = 0;
        if (sectionIndex == -1)
        {
            foreach (Section section in doc.Sections)
            {
                remainingCount += section.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).Count();
            }
        }
        else
        {
            var section = doc.Sections[sectionIndex];
            remainingCount = section.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Where(s => s.HasImage).Count();
        }

        var result = $"成功刪除 {imageInfo}\n";
        result += $"文檔剩餘圖片數: {remainingCount}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

