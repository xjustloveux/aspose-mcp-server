using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;

namespace AsposeMcpServer.Tools;

public class WordAddListItemTool : IAsposeTool
{
    public string Description => "Add a list item to Word document using a list style (supports multi-level lists with proper auto-numbering)";

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
            text = new
            {
                type = "string",
                description = "List item text content"
            },
            styleName = new
            {
                type = "string",
                description = "Style name for the list item (e.g., '!標題4-數字', '!標題5-數字括弧'). The style should contain list formatting."
            },
            listLevel = new
            {
                type = "number",
                description = "List level (0-8, where 0 is top level). This will apply the appropriate indentation from the style. Default: 0"
            },
            applyStyleIndent = new
            {
                type = "boolean",
                description = "If true, uses the indentation defined in the style. If false, uses manual listLevel. Default: true (recommended)"
            }
        },
        required = new[] { "path", "text", "styleName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var styleName = arguments?["styleName"]?.GetValue<string>() ?? throw new ArgumentException("styleName is required");
        var listLevel = arguments?["listLevel"]?.GetValue<int>() ?? 0;
        var applyStyleIndent = arguments?["applyStyleIndent"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        // Move to end of document
        builder.MoveToDocumentEnd();

        // Check if style exists
        var style = doc.Styles[styleName];
        if (style == null)
        {
            throw new ArgumentException($"找不到樣式 '{styleName}'，可用樣式請使用 word_get_styles 工具查看");
        }

        // Create paragraph with list item
        var para = new Paragraph(doc);
        
        // Apply the style (this includes all formatting, including indentation)
        para.ParagraphFormat.StyleName = styleName;

        // If the style doesn't naturally apply correct indentation for the level,
        // and user wants to use listLevel parameter
        if (!applyStyleIndent && listLevel > 0)
        {
            // Each level = 36 points (0.5 inch)
            para.ParagraphFormat.LeftIndent = listLevel * 36;
        }

        var run = new Run(doc, text);
        para.AppendChild(run);
        
        builder.CurrentParagraph.ParentNode.AppendChild(para);

        doc.Save(outputPath);

        var result = "成功添加清單項目\n";
        result += $"樣式: {styleName}\n";
        result += $"級別: {listLevel}\n";
        
        if (applyStyleIndent)
        {
            result += "縮排: 使用樣式定義的縮排（推薦）\n";
        }
        else if (listLevel > 0)
        {
            result += $"縮排: 手動設定 ({listLevel * 36} points)\n";
        }
        
        // Get the actual indent applied
        if (style.Type == StyleType.Paragraph)
        {
            var actualIndent = style.ParagraphFormat.LeftIndent;
            if (actualIndent > 0)
            {
                result += $"實際縮排: {actualIndent} points\n";
            }
        }
        
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

