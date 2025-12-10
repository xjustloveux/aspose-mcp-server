using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordAddTextWithStyleTool : IAsposeTool
{
    public string Description => "Add text to Word document with custom style (by style name or by defining style properties)";

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
                description = "Text content to add"
            },
            styleName = new
            {
                type = "string",
                description = "Style name to apply (e.g., 'Heading 1', '標題1', 'Normal'). If provided, other style properties are ignored."
            },
            fontName = new
            {
                type = "string",
                description = "Font name (e.g., '微軟雅黑', 'Arial'). If fontNameAscii and fontNameFarEast are provided, this will be used as fallback."
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
            underline = new
            {
                type = "boolean",
                description = "Underline text"
            },
            color = new
            {
                type = "string",
                description = "Text color (hex format like 'FF0000' for red, or name like 'Red', 'Blue')"
            },
            alignment = new
            {
                type = "string",
                description = "Text alignment: left, center, right, justify",
                @enum = new[] { "left", "center", "right", "justify" }
            },
            indentLevel = new
            {
                type = "number",
                description = "Indentation level (0-8, where each level = 36 points / 0.5 inch). 0 = no indent, 1 = first level, 2 = second level, etc."
            },
            leftIndent = new
            {
                type = "number",
                description = "Left indentation in points (alternative to indentLevel for precise control)"
            },
            firstLineIndent = new
            {
                type = "number",
                description = "First line indentation in points (positive = indent first line, negative = hanging indent)"
            },
            tabStops = new
            {
                type = "array",
                description = "Custom tab stops for this paragraph. Example: [{\"position\": 100, \"alignment\": \"Left\"}, {\"position\": 200, \"alignment\": \"Center\"}]",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        position = new { type = "number", description = "Tab stop position in points" },
                        alignment = new { type = "string", description = "Tab alignment: Left, Center, Right, Decimal, Bar", @enum = new[] { "Left", "Center", "Right", "Decimal", "Bar" } },
                        leader = new { type = "string", description = "Tab leader: None, Dots, Dashes, Line, Heavy, MiddleDot (default: None)", @enum = new[] { "None", "Dots", "Dashes", "Line", "Heavy", "MiddleDot" } }
                    }
                }
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Index of the paragraph to insert after (0-based). If not provided, text will be added at the end of document. Use -1 to insert at the beginning."
            }
        },
        required = new[] { "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var styleName = arguments?["styleName"]?.GetValue<string>();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var underline = arguments?["underline"]?.GetValue<bool?>();
        var color = arguments?["color"]?.GetValue<string>();
        var alignment = arguments?["alignment"]?.GetValue<string>();
        var indentLevel = arguments?["indentLevel"]?.GetValue<int?>();
        var leftIndent = arguments?["leftIndent"]?.GetValue<double?>();
        var firstLineIndent = arguments?["firstLineIndent"]?.GetValue<double?>();
        var tabStops = arguments?["tabStops"]?.AsArray();
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        Paragraph? targetPara = null;
        
        // Determine insertion position
        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                // Insert at the beginning - move to first paragraph
                if (paragraphs.Count > 0)
                {
                    targetPara = paragraphs[0] as Paragraph;
                    if (targetPara != null)
                    {
                        builder.MoveTo(targetPara);
                    }
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                // Insert after the specified paragraph
                targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                if (targetPara != null)
                {
                    builder.MoveTo(targetPara);
                }
                else
                {
                    throw new InvalidOperationException($"無法找到索引 {paragraphIndex.Value} 的段落");
                }
            }
            else
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
        }
        else
        {
            // Default: Move to end of document
            builder.MoveToDocumentEnd();
        }

        // Create paragraph
        var para = new Paragraph(doc);
        var run = new Run(doc, text);

        // Check if user is trying to override style properties (Bug 2.2 scenario)
        bool hasCustomParams = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) || 
                               !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue || 
                               bold.HasValue || italic.HasValue || underline.HasValue || 
                               !string.IsNullOrEmpty(color) || !string.IsNullOrEmpty(alignment);
        
        string warningMessage = "";
        if (!string.IsNullOrEmpty(styleName) && hasCustomParams)
        {
            warningMessage = "\n⚠️ 注意: 同時使用 styleName 和自訂參數時，自訂參數會覆蓋樣式中對應的屬性。\n" +
                           "這允許您在應用樣式的同時自訂特定屬性。\n" +
                           "如果需要完全自訂的樣式，建議使用 word_create_style 創建自訂樣式。\n" +
                           "範例: word_create_style(styleName='自訂標題', baseStyle='Heading 1', color='000000')";
        }

        // Apply style by name if provided
        if (!string.IsNullOrEmpty(styleName))
        {
            try
            {
                // Check if style exists
                var style = doc.Styles[styleName];
                if (style != null)
                {
                    para.ParagraphFormat.StyleName = styleName;
                }
                else
                {
                    throw new ArgumentException($"找不到樣式 '{styleName}'，可用樣式請使用 word_get_styles 工具查看");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法應用樣式 '{styleName}': {ex.Message}，可用樣式請使用 word_get_styles 工具查看", ex);
            }
        }
        
        // Apply custom formatting parameters (these override style defaults)
        // This allows users to use a style but customize specific properties
        
        // Set font names (priority: fontNameAscii/fontNameFarEast > fontName)
        if (!string.IsNullOrEmpty(fontNameAscii))
            run.Font.NameAscii = fontNameAscii;
        
        if (!string.IsNullOrEmpty(fontNameFarEast))
            run.Font.NameFarEast = fontNameFarEast;
        
        if (!string.IsNullOrEmpty(fontName))
        {
            // If fontNameAscii/FarEast are not set, use fontName for both
            if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
            {
                run.Font.Name = fontName;
            }
            else
            {
                // If only one is set, use fontName as fallback for the other
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
        
        if (underline.HasValue)
            run.Font.Underline = underline.Value ? Underline.Single : Underline.None;
        
        if (!string.IsNullOrEmpty(color))
        {
            try
            {
                // Try to parse as hex color
                if (color.StartsWith("#"))
                    color = color.Substring(1);
                
                if (color.Length == 6)
                {
                    int r = Convert.ToInt32(color.Substring(0, 2), 16);
                    int g = Convert.ToInt32(color.Substring(2, 2), 16);
                    int b = Convert.ToInt32(color.Substring(4, 2), 16);
                    run.Font.Color = System.Drawing.Color.FromArgb(r, g, b);
                }
                else
                {
                    // Try to parse as color name
                    run.Font.Color = System.Drawing.Color.FromName(color);
                }
            }
            catch
            {
                // Ignore color parsing errors
            }
        }

        if (!string.IsNullOrEmpty(alignment))
        {
            para.ParagraphFormat.Alignment = alignment.ToLower() switch
            {
                "left" => ParagraphAlignment.Left,
                "right" => ParagraphAlignment.Right,
                "center" => ParagraphAlignment.Center,
                "justify" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left
            };
        }

        // Apply indentation (only if not already set by style)
        // Note: If the style has indentation (e.g., list styles), it will be preserved
        // Manual indentation parameters can override if explicitly provided
        if (indentLevel.HasValue)
        {
            // Each level = 36 points (0.5 inch)
            para.ParagraphFormat.LeftIndent = indentLevel.Value * 36;
        }
        else if (leftIndent.HasValue)
        {
            para.ParagraphFormat.LeftIndent = leftIndent.Value;
        }
        // If neither indentLevel nor leftIndent is specified, keep the style's indent setting

        if (firstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
        }
        // If firstLineIndent not specified, keep the style's firstLineIndent setting

        // Apply tab stops
        if (tabStops != null && tabStops.Count > 0)
        {
            para.ParagraphFormat.TabStops.Clear(); // Clear existing tab stops first
            
            foreach (var tabStopJson in tabStops)
            {
                var position = tabStopJson?["position"]?.GetValue<double>() ?? 0;
                var alignmentStr = tabStopJson?["alignment"]?.GetValue<string>() ?? "Left";
                var leaderStr = tabStopJson?["leader"]?.GetValue<string>() ?? "None";
                
                var tabAlignment = alignmentStr switch
                {
                    "Center" => TabAlignment.Center,
                    "Right" => TabAlignment.Right,
                    "Decimal" => TabAlignment.Decimal,
                    "Bar" => TabAlignment.Bar,
                    _ => TabAlignment.Left
                };
                
                var tabLeader = leaderStr switch
                {
                    "Dots" => TabLeader.Dots,
                    "Dashes" => TabLeader.Dashes,
                    "Line" => TabLeader.Line,
                    "Heavy" => TabLeader.Heavy,
                    "MiddleDot" => TabLeader.MiddleDot,
                    _ => TabLeader.None
                };
                
                para.ParagraphFormat.TabStops.Add(new TabStop(position, tabAlignment, tabLeader));
            }
        }

        para.AppendChild(run);
        
        // Insert paragraph at the correct position
        if (paragraphIndex.HasValue && targetPara != null)
        {
            if (paragraphIndex.Value == -1)
            {
                // Insert at the beginning - insert before first paragraph
                targetPara.ParentNode.InsertBefore(para, targetPara);
            }
            else
            {
                // Insert after the specified paragraph
                targetPara.ParentNode.InsertAfter(para, targetPara);
            }
        }
        else
        {
            // Default: Append to end using builder
            builder.CurrentParagraph.ParentNode.AppendChild(para);
        }

        doc.Save(outputPath);

        var result = "成功添加文本\n";
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                result += "插入位置: 文檔開頭\n";
            }
            else
            {
                result += $"插入位置: 段落 #{paragraphIndex.Value} 之後\n";
            }
        }
        else
        {
            result += "插入位置: 文檔末尾\n";
        }
        
        if (!string.IsNullOrEmpty(styleName))
        {
            result += $"應用樣式: {styleName}\n";
        }
        else
        {
            result += "自定義格式:\n";
            if (!string.IsNullOrEmpty(fontNameAscii)) result += $"  字體（英文）: {fontNameAscii}\n";
            if (!string.IsNullOrEmpty(fontNameFarEast)) result += $"  字體（中文）: {fontNameFarEast}\n";
            if (!string.IsNullOrEmpty(fontName) && string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast)) 
                result += $"  字體: {fontName}\n";
            if (fontSize.HasValue) result += $"  字號: {fontSize.Value} pt\n";
            if (bold.HasValue && bold.Value) result += $"  粗體\n";
            if (italic.HasValue && italic.Value) result += $"  斜體\n";
            if (underline.HasValue && underline.Value) result += $"  底線\n";
            if (!string.IsNullOrEmpty(color)) result += $"  顏色: {color}\n";
            if (!string.IsNullOrEmpty(alignment)) result += $"  對齊: {alignment}\n";
        }
        if (indentLevel.HasValue) result += $"縮排級別: {indentLevel.Value} ({indentLevel.Value * 36} pt)\n";
        else if (leftIndent.HasValue) result += $"左縮排: {leftIndent.Value} pt\n";
        if (firstLineIndent.HasValue) result += $"首行縮排: {firstLineIndent.Value} pt\n";
        result += $"輸出: {outputPath}";
        
        // Add warning message if applicable
        result += warningMessage;

        return await Task.FromResult(result);
    }
}

