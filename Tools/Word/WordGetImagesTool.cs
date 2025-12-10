using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordGetImagesTool : IAsposeTool
{
    public string Description => "Get all images information from a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

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
}

