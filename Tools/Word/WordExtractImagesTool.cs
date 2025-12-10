using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordExtractImagesTool : IAsposeTool
{
    public string Description => "Extract all images from a Word document to a specified directory";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input Word document path"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory for extracted images"
            },
            prefix = new
            {
                type = "string",
                description = "Filename prefix for images (default: 'image')"
            }
        },
        required = new[] { "path", "outputDir" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? throw new ArgumentException("outputDir is required");
        var prefix = arguments?["prefix"]?.GetValue<string>() ?? "image";

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
            
            // Determine file extension based on image bytes
            var imageBytes = imageData.ImageBytes;
            string extension = "img";
            
            if (imageBytes != null && imageBytes.Length > 4)
            {
                // Check file signature
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
            
            // Save image
            using (var stream = File.Create(outputPath))
            {
                imageData.Save(stream);
            }
            
            extractedFiles.Add(outputPath);
        }

        return await Task.FromResult($"成功提取 {shapes.Count} 張圖片到: {outputDir}\n" +
                                    $"文件列表:\n" + string.Join("\n", extractedFiles.Select(f => $"  - {Path.GetFileName(f)}")));
    }
}

