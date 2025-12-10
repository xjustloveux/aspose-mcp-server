using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.SmartArt;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptManageSmartArtNodesTool : IAsposeTool
{
    public string Description => "Add/remove/rename/move SmartArt nodes by path";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndex = new { type = "number", description = "Slide index (0-based)" },
            shapeIndex = new { type = "number", description = "SmartArt shape index (0-based)" },
            action = new { type = "string", description = "add|remove|rename|move" },
            targetPath = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Node path indices from root (e.g., [0,1])"
            },
            text = new { type = "string", description = "Text for add/rename" },
            moveParentPath = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Destination parent path for move"
            },
            moveIndex = new { type = "number", description = "Insert index under new parent (optional, default append)" }
        },
        required = new[] { "path", "slideIndex", "shapeIndex", "action" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var action = arguments?["action"]?.GetValue<string>()?.ToLower() ?? throw new ArgumentException("action is required");
        var targetPath = arguments?["targetPath"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var text = arguments?["text"]?.GetValue<string>();
        var moveParentPath = arguments?["moveParentPath"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var moveIndex = arguments?["moveIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }
        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
        }

        if (slide.Shapes[shapeIndex] is not ISmartArt smartArt)
        {
            throw new ArgumentException("指定的 shape 不是 SmartArt");
        }

        SmartArtNode? GetNode(int[]? pathArr)
        {
            if (pathArr == null || pathArr.Length == 0) return null;
            var node = smartArt.AllNodes[pathArr[0]] as SmartArtNode;
            for (int i = 1; node != null && i < pathArr.Length; i++)
            {
                if (pathArr[i] < 0 || pathArr[i] >= node.ChildNodes.Count) return null;
                node = node.ChildNodes[pathArr[i]] as SmartArtNode;
            }
            return node;
        }

        switch (action)
        {
            case "add":
                {
                    var parent = GetNode(targetPath) ?? smartArt.AllNodes.AddNode();
                    var newNode = parent.ChildNodes.AddNode();
                    newNode.TextFrame.Text = string.IsNullOrEmpty(text) ? "New Node" : text;
                    break;
                }
            case "remove":
                {
                    var node = GetNode(targetPath);
                    if (node == null) throw new ArgumentException("targetPath 無效");
                    node.Remove();
                    break;
                }
            case "rename":
                {
                    var node = GetNode(targetPath);
                    if (node == null) throw new ArgumentException("targetPath 無效");
                    node.TextFrame.Text = text ?? string.Empty;
                    break;
                }
            case "move":
                {
                    var node = GetNode(targetPath);
                    if (node == null) throw new ArgumentException("targetPath 無效");
                    var parent = GetNode(moveParentPath) ?? smartArt.AllNodes.AddNode();
                    var clone = parent.ChildNodes.AddNode();
                    clone.TextFrame.Text = node.TextFrame.Text;
                    foreach (SmartArtNode child in node.ChildNodes)
                    {
                        var childClone = clone.ChildNodes.AddNode();
                        childClone.TextFrame.Text = child.TextFrame.Text;
                    }
                    node.Remove();
                    break;
                }
            default:
                throw new ArgumentException("action 必須為 add/remove/rename/move");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"SmartArt {action} 完成：slide {slideIndex}, shape {shapeIndex}");
    }
}

