using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordGetTextboxesTool : IAsposeTool
{
    public string Description => "Get all textboxes from a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            includeContent = new
            {
                type = "boolean",
                description = "Include textbox content (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var includeContent = arguments?["includeContent"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
            .Where(s => s.ShapeType == ShapeType.TextBox)
            .ToList();
        
        var result = new StringBuilder();

        result.AppendLine("=== 文檔文本框資訊 ===\n");
        result.AppendLine($"總文本框數: {shapes.Count}\n");

        if (shapes.Count == 0)
        {
            result.AppendLine("未找到文本框");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < shapes.Count; i++)
        {
            var textbox = shapes[i];
            result.AppendLine($"【文本框 {i}】");
            result.AppendLine($"名稱: {textbox.Name ?? "(無名稱)"}");
            result.AppendLine($"寬度: {textbox.Width} 點");
            result.AppendLine($"高度: {textbox.Height} 點");
            result.AppendLine($"位置: X={textbox.Left}, Y={textbox.Top}");
            // Note: Shape.Locked and Shape.Visible properties may not be available in this API version
            
            if (includeContent)
            {
                var textboxText = textbox.GetText().Trim();
                if (!string.IsNullOrEmpty(textboxText))
                {
                    result.AppendLine($"內容:");
                    result.AppendLine($"  {textboxText.Replace("\n", "\n  ")}");
                }
                else
                {
                    result.AppendLine($"內容: (空)");
                }
            }
            
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

