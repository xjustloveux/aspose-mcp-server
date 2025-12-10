using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Text;

namespace AsposeMcpServer.Tools;

public class PptReplaceTextTool : IAsposeTool
{
    public string Description => "Find and replace text across PowerPoint slides";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            findText = new
            {
                type = "string",
                description = "Text to find"
            },
            replaceText = new
            {
                type = "string",
                description = "Text to replace with"
            },
            matchCase = new
            {
                type = "boolean",
                description = "Match case (default: false)"
            }
        },
        required = new[] { "path", "findText", "replaceText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var findText = arguments?["findText"]?.GetValue<string>() ?? throw new ArgumentException("findText is required");
        var replaceText = arguments?["replaceText"]?.GetValue<string>() ?? throw new ArgumentException("replaceText is required");
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

