using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetStylesTool : IAsposeTool
{
    public string Description => "Get all styles defined in a Word document";

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
            includeBuiltIn = new
            {
                type = "boolean",
                description = "Include built-in styles (default: false)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var includeBuiltIn = arguments?["includeBuiltIn"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        var result = new StringBuilder();
        
        result.AppendLine("=== 文檔樣式信息 ===\n");

        // Paragraph styles
        result.AppendLine("【段落樣式】");
        var paraStyles = doc.Styles
            .Cast<Style>()
            .Where(s => s.Type == StyleType.Paragraph && (includeBuiltIn || !s.BuiltIn))
            .OrderBy(s => s.Name)
            .ToList();

        foreach (var style in paraStyles)
        {
            result.AppendLine($"\n樣式名稱: {style.Name}");
            if (style.BaseStyleName != null && !string.IsNullOrEmpty(style.BaseStyleName))
            {
                result.AppendLine($"  基於: {style.BaseStyleName}");
            }
            
            var font = style.Font;
            
            // Show separate fonts for ASCII (English) and Far East (Chinese/Japanese/Korean)
            if (font.NameAscii != font.NameFarEast)
            {
                result.AppendLine($"  字體（英文）: {font.NameAscii}");
                result.AppendLine($"  字體（中文）: {font.NameFarEast}");
            }
            else
            {
                result.AppendLine($"  字體: {font.Name}");
            }
            
            result.AppendLine($"  字號: {font.Size} pt");
            
            if (font.Bold) result.AppendLine($"  樣式: 粗體");
            if (font.Italic) result.AppendLine($"  樣式: 斜體");
            if (font.Underline != Underline.None) result.AppendLine($"  樣式: 底線 ({font.Underline})");
            
            if (font.Color.ToArgb() != System.Drawing.Color.Empty.ToArgb())
            {
                result.AppendLine($"  顏色: #{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}");
            }

            var paraFormat = style.ParagraphFormat;
            result.AppendLine($"  對齊: {paraFormat.Alignment}");
            
            // Show paragraph spacing
            if (paraFormat.SpaceBefore != 0)
                result.AppendLine($"  段落前間距: {paraFormat.SpaceBefore} pt");
            if (paraFormat.SpaceAfter != 0)
                result.AppendLine($"  段落後間距: {paraFormat.SpaceAfter} pt");
            
            result.AppendLine($"  行距: {paraFormat.LineSpacing} pt ({paraFormat.LineSpacingRule})");
            
            if (paraFormat.LeftIndent != 0)
                result.AppendLine($"  左縮進: {paraFormat.LeftIndent} pt");
            if (paraFormat.RightIndent != 0)
                result.AppendLine($"  右縮進: {paraFormat.RightIndent} pt");
            if (paraFormat.FirstLineIndent != 0)
                result.AppendLine($"  首行縮進: {paraFormat.FirstLineIndent} pt");
        }

        // Character styles
        result.AppendLine("\n\n【字符樣式】");
        var charStyles = doc.Styles
            .Cast<Style>()
            .Where(s => s.Type == StyleType.Character && (includeBuiltIn || !s.BuiltIn))
            .OrderBy(s => s.Name)
            .ToList();

        foreach (var style in charStyles)
        {
            result.AppendLine($"\n樣式名稱: {style.Name}");
            var font = style.Font;
            
            // Show separate fonts for ASCII (English) and Far East (Chinese/Japanese/Korean)
            if (font.NameAscii != font.NameFarEast)
            {
                result.AppendLine($"  字體（英文）: {font.NameAscii}");
                result.AppendLine($"  字體（中文）: {font.NameFarEast}");
            }
            else
            {
                result.AppendLine($"  字體: {font.Name}");
            }
            
            result.AppendLine($"  字號: {font.Size} pt");
            
            if (font.Bold) result.AppendLine($"  樣式: 粗體");
            if (font.Italic) result.AppendLine($"  樣式: 斜體");
            if (font.Underline != Underline.None) result.AppendLine($"  樣式: 底線");
            
            if (font.Color.ToArgb() != System.Drawing.Color.Empty.ToArgb())
            {
                result.AppendLine($"  顏色: #{font.Color.R:X2}{font.Color.G:X2}{font.Color.B:X2}");
            }
        }

        // Table styles
        result.AppendLine("\n\n【表格樣式】");
        var tableStyles = doc.Styles
            .Cast<Style>()
            .Where(s => s.Type == StyleType.Table && (includeBuiltIn || !s.BuiltIn))
            .OrderBy(s => s.Name)
            .ToList();

        foreach (var style in tableStyles)
        {
            result.AppendLine($"\n樣式名稱: {style.Name}");
            if (style.BaseStyleName != null && !string.IsNullOrEmpty(style.BaseStyleName))
            {
                result.AppendLine($"  基於: {style.BaseStyleName}");
            }
        }

        result.AppendLine($"\n\n總計:");
        result.AppendLine($"  段落樣式: {paraStyles.Count}");
        result.AppendLine($"  字符樣式: {charStyles.Count}");
        result.AppendLine($"  表格樣式: {tableStyles.Count}");

        return await Task.FromResult(result.ToString());
    }
}

