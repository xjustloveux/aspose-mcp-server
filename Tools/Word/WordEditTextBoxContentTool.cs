using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordEditTextBoxContentTool : IAsposeTool
{
    public string Description => "Edit the content and formatting of an existing text box in a Word document";

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
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            textboxIndex = new
            {
                type = "number",
                description = "Text box index (0-based)"
            },
            text = new
            {
                type = "string",
                description = "New text content (if not provided, keeps existing text)"
            },
            appendText = new
            {
                type = "boolean",
                description = "Append text to existing content instead of replacing (default: false)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name. If fontNameAscii and fontNameFarEast are provided, this will be used as fallback."
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, e.g., 'Times New Roman')"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean, e.g., '標楷體')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text"
            },
            color = new
            {
                type = "string",
                description = "Text color (hex format like '#FF0000' or 'FF0000' for red)"
            },
            alignment = new
            {
                type = "string",
                description = "Text alignment: left, center, right",
                @enum = new[] { "left", "center", "right" }
            },
            clearFormatting = new
            {
                type = "boolean",
                description = "Clear existing formatting before applying new formatting (default: false)"
            }
        },
        required = new[] { "path", "textboxIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var textboxIndex = arguments?["textboxIndex"]?.GetValue<int>() ?? throw new ArgumentException("textboxIndex is required");
        var text = arguments?["text"]?.GetValue<string>();
        var appendText = arguments?["appendText"]?.GetValue<bool>() ?? false;
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var color = arguments?["color"]?.GetValue<string>();
        var alignment = arguments?["alignment"]?.GetValue<string>();
        var clearFormatting = arguments?["clearFormatting"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        
        // Get all text boxes
        var shapes = doc.GetChildNodes(NodeType.Shape, true);
        var textboxes = shapes.Cast<Shape>().Where(s => s.ShapeType == ShapeType.TextBox).ToList();
        
        if (textboxIndex < 0 || textboxIndex >= textboxes.Count)
        {
            throw new ArgumentException($"文字框索引 {textboxIndex} 超出範圍 (文檔共有 {textboxes.Count} 個文字框)");
        }
        
        var textbox = textboxes[textboxIndex];
        
        // Get or create paragraph in textbox
        var paragraphs = textbox.GetChildNodes(NodeType.Paragraph, true);
        Paragraph para;
        
        if (paragraphs.Count == 0)
        {
            para = new Paragraph(doc);
            textbox.AppendChild(para);
        }
        else
        {
            para = paragraphs[0] as Paragraph ?? throw new Exception("無法取得文字框段落");
        }
        
        var changes = new List<string>();
        
        // Update text content
        if (text != null)
        {
            if (appendText && para.Runs.Count > 0)
            {
                // Append to existing text
                var newRun = new Run(doc, text);
                para.AppendChild(newRun);
                changes.Add($"附加文字: {text.Substring(0, Math.Min(50, text.Length))}{(text.Length > 50 ? "..." : "")}");
            }
            else
            {
                // Replace all text
                para.RemoveAllChildren();
                var newRun = new Run(doc, text);
                para.AppendChild(newRun);
                changes.Add($"設定文字: {text.Substring(0, Math.Min(50, text.Length))}{(text.Length > 50 ? "..." : "")}");
            }
        }
        
        // Apply formatting to all runs in the textbox
        var runs = para.GetChildNodes(NodeType.Run, false);
        
        if (clearFormatting)
        {
            foreach (Run run in runs)
            {
                run.Font.ClearFormatting();
            }
            changes.Add("清除現有格式");
        }
        
        bool hasFormatting = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) || 
                             !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue || 
                             bold.HasValue || italic.HasValue || !string.IsNullOrEmpty(color);
        
        if (hasFormatting)
        {
            foreach (Run run in runs)
            {
                // Set font names (priority: fontNameAscii/fontNameFarEast > fontName)
                if (!string.IsNullOrEmpty(fontNameAscii))
                    run.Font.NameAscii = fontNameAscii;
                
                if (!string.IsNullOrEmpty(fontNameFarEast))
                    run.Font.NameFarEast = fontNameFarEast;
                
                if (!string.IsNullOrEmpty(fontName))
                {
                    if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                    {
                        run.Font.Name = fontName;
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(fontNameAscii))
                            run.Font.NameAscii = fontName;
                        if (string.IsNullOrEmpty(fontNameFarEast))
                            run.Font.NameFarEast = fontName;
                    }
                }
                
                if (fontSize.HasValue)
                    run.Font.Size = fontSize.Value;
                
                if (bold.HasValue)
                    run.Font.Bold = bold.Value;
                
                if (italic.HasValue)
                    run.Font.Italic = italic.Value;
                
                if (!string.IsNullOrEmpty(color))
                {
                    try
                    {
                        var colorStr = color.TrimStart('#');
                        if (colorStr.Length == 6)
                        {
                            int r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
                            int g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
                            int b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
                            run.Font.Color = System.Drawing.Color.FromArgb(r, g, b);
                        }
                    }
                    catch { }
                }
            }
            
            var formatList = new List<string>();
            if (!string.IsNullOrEmpty(fontNameAscii)) formatList.Add($"字體（英文）: {fontNameAscii}");
            if (!string.IsNullOrEmpty(fontNameFarEast)) formatList.Add($"字體（中文）: {fontNameFarEast}");
            if (!string.IsNullOrEmpty(fontName) && string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast)) 
                formatList.Add($"字體: {fontName}");
            if (fontSize.HasValue) formatList.Add($"字號: {fontSize.Value}pt");
            if (bold.HasValue && bold.Value) formatList.Add("粗體");
            if (italic.HasValue && italic.Value) formatList.Add("斜體");
            if (!string.IsNullOrEmpty(color)) formatList.Add($"顏色: {color}");
            
            if (formatList.Count > 0)
                changes.Add($"文字格式: {string.Join(", ", formatList)}");
        }
        
        // Apply alignment
        if (!string.IsNullOrEmpty(alignment))
        {
            para.ParagraphFormat.Alignment = alignment.ToLower() switch
            {
                "center" => ParagraphAlignment.Center,
                "right" => ParagraphAlignment.Right,
                _ => ParagraphAlignment.Left
            };
            changes.Add($"對齊方式: {alignment}");
        }
        
        doc.Save(outputPath);
        
        var result = $"成功編輯文字框 #{textboxIndex}\n";
        if (changes.Count > 0)
        {
            result += $"變更內容:\n";
            foreach (var change in changes)
            {
                result += $"  - {change}\n";
            }
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

