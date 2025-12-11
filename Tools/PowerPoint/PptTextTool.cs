using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint text (add, edit, replace)
/// Merges: PptAddTextTool, PptEditTextTool, PptReplaceTextTool
/// </summary>
public class PptTextTool : IAsposeTool
{
    public string Description => "Manage PowerPoint text: add, edit, or replace";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'edit', 'replace'",
                @enum = new[] { "add", "edit", "replace" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for add/edit)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based, required for edit)"
            },
            text = new
            {
                type = "string",
                description = "Text content (required for add/edit)"
            },
            findText = new
            {
                type = "string",
                description = "Text to find (required for replace)"
            },
            replaceText = new
            {
                type = "string",
                description = "Text to replace with (required for replace)"
            },
            matchCase = new
            {
                type = "boolean",
                description = "Match case (optional, for replace, default: false)"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, for add, default: 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, for add, default: 50)"
            },
            width = new
            {
                type = "number",
                description = "Text box width (optional, for add, default: 400)"
            },
            height = new
            {
                type = "number",
                description = "Text box height (optional, for add, default: 100)"
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
            "add" => await AddTextAsync(arguments, path),
            "edit" => await EditTextAsync(arguments, path),
            "replace" => await ReplaceTextAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddTextAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for add operation");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required for add operation");
        var x = arguments?["x"]?.GetValue<float>() ?? 50;
        var y = arguments?["y"]?.GetValue<float>() ?? 50;
        var width = arguments?["width"]?.GetValue<float>() ?? 400;
        var height = arguments?["height"]?.GetValue<float>() ?? 100;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
        textBox.TextFrame.Text = text;

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Text added to slide {slideIndex}: {path}");
    }

    private async Task<string> EditTextAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for edit operation");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for edit operation");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required for edit operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
        {
            autoShape.TextFrame.Text = text;
        }
        else
        {
            throw new ArgumentException($"Shape at index {shapeIndex} does not support text editing");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Text updated on slide {slideIndex}, shape {shapeIndex}");
    }

    private async Task<string> ReplaceTextAsync(JsonObject? arguments, string path)
    {
        var findText = arguments?["findText"]?.GetValue<string>() ?? throw new ArgumentException("findText is required for replace operation");
        var replaceText = arguments?["replaceText"]?.GetValue<string>() ?? throw new ArgumentException("replaceText is required for replace operation");
        var matchCase = arguments?["matchCase"]?.GetValue<bool>() ?? false;

        using var presentation = new Presentation(path);
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var replacements = 0;

        foreach (var slide in presentation.Slides)
        {
            foreach (var shape in slide.Shapes)
            {
                if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
                {
                    foreach (var paragraph in autoShape.TextFrame.Paragraphs)
                    {
                        foreach (var portion in paragraph.Portions)
                        {
                            var text = portion.Text;
                            if (string.IsNullOrEmpty(text)) continue;

                            var newText = ReplaceAll(text, findText, replaceText, comparison);
                            if (!ReferenceEquals(text, newText) && newText != text)
                            {
                                portion.Text = newText;
                                replacements++;
                            }
                        }
                    }
                }
            }
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"完成文字替換：{replacements} 個片段\n查找: {findText}\n替換為: {replaceText}\n輸出: {path}");
    }

    private static string ReplaceAll(string source, string find, string replace, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(find)) return source;

        var sb = new StringBuilder();
        var idx = 0;
        while (true)
        {
            var next = source.IndexOf(find, idx, comparison);
            if (next < 0)
            {
                sb.Append(source, idx, source.Length - idx);
                break;
            }

            sb.Append(source, idx, next - idx);
            sb.Append(replace);
            idx = next + find.Length;
        }

        return sb.ToString();
    }
}

