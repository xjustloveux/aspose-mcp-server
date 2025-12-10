using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordEditParagraphTool : IAsposeTool
{
    public string Description => "Edit formatting of an existing paragraph in a Word document";

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
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            },
            // Font properties
            fontName = new
            {
                type = "string",
                description = "Font name (e.g., '標楷體', 'Arial')"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean)"
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
                description = "Font color (hex format, e.g., '000000' for black)"
            },
            // Paragraph properties
            alignment = new
            {
                type = "string",
                description = "Paragraph alignment: left, center, right, justify",
                @enum = new[] { "left", "center", "right", "justify" }
            },
            indentLeft = new
            {
                type = "number",
                description = "Left indent in points"
            },
            indentRight = new
            {
                type = "number",
                description = "Right indent in points"
            },
            firstLineIndent = new
            {
                type = "number",
                description = "First line indent in points (positive for indent, negative for hanging)"
            },
            spaceBefore = new
            {
                type = "number",
                description = "Space before paragraph in points"
            },
            spaceAfter = new
            {
                type = "number",
                description = "Space after paragraph in points"
            },
            lineSpacing = new
            {
                type = "number",
                description = "Line spacing (points or multiplier depending on lineSpacingRule)"
            },
            lineSpacingRule = new
            {
                type = "string",
                description = "Line spacing rule: single, oneAndHalf, double, atLeast, exactly, multiple",
                @enum = new[] { "single", "oneAndHalf", "double", "atLeast", "exactly", "multiple" }
            },
            styleName = new
            {
                type = "string",
                description = "Style name to apply to paragraph"
            },
            tabStops = new
            {
                type = "array",
                description = "Custom tab stops (array of objects with position, alignment, leader)",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        position = new { type = "number" },
                        alignment = new { type = "string", @enum = new[] { "left", "center", "right", "decimal", "bar", "clear" } },
                        leader = new { type = "string", @enum = new[] { "none", "dots", "dashes", "line", "heavy", "middleDot" } }
                    }
                }
            }
        },
        required = new[] { "path", "paragraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
        
        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        
        if (paragraphIndex >= paragraphs.Count)
            throw new ArgumentException($"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count})");
        
        var para = paragraphs[paragraphIndex];
        var builder = new DocumentBuilder(doc);
        builder.MoveTo(para.FirstChild);
        
        // Apply font properties
        if (arguments?["fontName"] != null)
        {
            var fontName = arguments["fontName"]?.GetValue<string>();
            builder.Font.Name = fontName ?? "";
        }
        
        if (arguments?["fontNameAscii"] != null)
        {
            var fontNameAscii = arguments["fontNameAscii"]?.GetValue<string>();
            builder.Font.NameAscii = fontNameAscii ?? "";
        }
        
        if (arguments?["fontNameFarEast"] != null)
        {
            var fontNameFarEast = arguments["fontNameFarEast"]?.GetValue<string>();
            builder.Font.NameFarEast = fontNameFarEast ?? "";
        }
        
        if (arguments?["fontSize"] != null)
        {
            var fontSize = arguments["fontSize"]?.GetValue<double>();
            if (fontSize.HasValue)
                builder.Font.Size = fontSize.Value;
        }
        
        if (arguments?["bold"] != null)
        {
            builder.Font.Bold = arguments["bold"]?.GetValue<bool>() ?? false;
        }
        
        if (arguments?["italic"] != null)
        {
            builder.Font.Italic = arguments["italic"]?.GetValue<bool>() ?? false;
        }
        
        if (arguments?["underline"] != null)
        {
            var underline = arguments["underline"]?.GetValue<bool>() ?? false;
            builder.Font.Underline = underline ? Underline.Single : Underline.None;
        }
        
        if (arguments?["color"] != null)
        {
            var colorStr = arguments["color"]?.GetValue<string>();
            if (!string.IsNullOrEmpty(colorStr))
            {
                builder.Font.Color = ParseColor(colorStr);
            }
        }
        
        // Apply paragraph properties
        var paraFormat = para.ParagraphFormat;
        
        if (arguments?["alignment"] != null)
        {
            var alignment = arguments["alignment"]?.GetValue<string>() ?? "left";
            paraFormat.Alignment = GetAlignment(alignment);
        }
        
        if (arguments?["indentLeft"] != null)
        {
            var indentLeft = arguments["indentLeft"]?.GetValue<double>();
            if (indentLeft.HasValue)
                paraFormat.LeftIndent = indentLeft.Value;
        }
        
        if (arguments?["indentRight"] != null)
        {
            var indentRight = arguments["indentRight"]?.GetValue<double>();
            if (indentRight.HasValue)
                paraFormat.RightIndent = indentRight.Value;
        }
        
        if (arguments?["firstLineIndent"] != null)
        {
            var firstLineIndent = arguments["firstLineIndent"]?.GetValue<double>();
            if (firstLineIndent.HasValue)
                paraFormat.FirstLineIndent = firstLineIndent.Value;
        }
        
        if (arguments?["spaceBefore"] != null)
        {
            var spaceBefore = arguments["spaceBefore"]?.GetValue<double>();
            if (spaceBefore.HasValue)
                paraFormat.SpaceBefore = spaceBefore.Value;
        }
        
        if (arguments?["spaceAfter"] != null)
        {
            var spaceAfter = arguments["spaceAfter"]?.GetValue<double>();
            if (spaceAfter.HasValue)
                paraFormat.SpaceAfter = spaceAfter.Value;
        }
        
        if (arguments?["lineSpacing"] != null || arguments?["lineSpacingRule"] != null)
        {
            var lineSpacing = arguments?["lineSpacing"]?.GetValue<double>();
            var lineSpacingRule = arguments?["lineSpacingRule"]?.GetValue<string>() ?? "single";
            
            var rule = GetLineSpacingRule(lineSpacingRule);
            paraFormat.LineSpacingRule = rule;
            
            if (lineSpacing.HasValue)
            {
                paraFormat.LineSpacing = lineSpacing.Value;
            }
            else if (lineSpacingRule == "single")
            {
                paraFormat.LineSpacing = 12; // Default single spacing
            }
            else if (lineSpacingRule == "oneAndHalf")
            {
                paraFormat.LineSpacing = 18; // 1.5x spacing
            }
            else if (lineSpacingRule == "double")
            {
                paraFormat.LineSpacing = 24; // Double spacing
            }
        }
        
        if (arguments?["styleName"] != null)
        {
            var styleName = arguments["styleName"]?.GetValue<string>();
            if (!string.IsNullOrEmpty(styleName))
            {
                try
                {
                    paraFormat.Style = doc.Styles[styleName];
                }
                catch
                {
                    // Style not found, ignore
                }
            }
        }
        
        // Apply tab stops
        if (arguments?["tabStops"] != null)
        {
            var tabStops = arguments["tabStops"]?.AsArray();
            if (tabStops != null && tabStops.Count > 0)
            {
                paraFormat.TabStops.Clear();
                foreach (var ts in tabStops)
                {
                    var position = ts?["position"]?.GetValue<double>() ?? 0;
                    var alignment = ts?["alignment"]?.GetValue<string>() ?? "left";
                    var leader = ts?["leader"]?.GetValue<string>() ?? "none";
                    
                    paraFormat.TabStops.Add(new TabStop(
                        position,
                        GetTabAlignment(alignment),
                        GetTabLeader(leader)
                    ));
                }
            }
        }
        
        // Apply font to all runs in paragraph
        foreach (Run run in para.GetChildNodes(NodeType.Run, true))
        {
            if (arguments?["fontName"] != null)
            {
                var fontName = arguments["fontName"]?.GetValue<string>();
                run.Font.Name = fontName ?? "";
            }
            
            if (arguments?["fontNameAscii"] != null)
            {
                var fontNameAscii = arguments["fontNameAscii"]?.GetValue<string>();
                run.Font.NameAscii = fontNameAscii ?? "";
            }
            
            if (arguments?["fontNameFarEast"] != null)
            {
                var fontNameFarEast = arguments["fontNameFarEast"]?.GetValue<string>();
                run.Font.NameFarEast = fontNameFarEast ?? "";
            }
            
            if (arguments?["fontSize"] != null)
            {
                var fontSize = arguments["fontSize"]?.GetValue<double>();
                if (fontSize.HasValue)
                    run.Font.Size = fontSize.Value;
            }
            
            if (arguments?["bold"] != null)
            {
                run.Font.Bold = arguments["bold"]?.GetValue<bool>() ?? false;
            }
            
            if (arguments?["italic"] != null)
            {
                run.Font.Italic = arguments["italic"]?.GetValue<bool>() ?? false;
            }
            
            if (arguments?["underline"] != null)
            {
                var underline = arguments["underline"]?.GetValue<bool>() ?? false;
                run.Font.Underline = underline ? Underline.Single : Underline.None;
            }
            
            if (arguments?["color"] != null)
            {
                var colorStr = arguments["color"]?.GetValue<string>();
                if (!string.IsNullOrEmpty(colorStr))
                {
                    run.Font.Color = ParseColor(colorStr);
                }
            }
        }
        
        doc.Save(outputPath);
        
        return await Task.FromResult($"成功編輯段落 {paragraphIndex} 的格式");
    }
    
    private ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" => ParagraphAlignment.Justify,
            _ => ParagraphAlignment.Left
        };
    }
    
    private LineSpacingRule GetLineSpacingRule(string rule)
    {
        return rule.ToLower() switch
        {
            "single" => LineSpacingRule.Exactly, // Use Exactly with LineSpacing = 12 for single
            "oneAndHalf" => LineSpacingRule.Exactly, // Use Exactly with LineSpacing = 18 for 1.5x
            "double" => LineSpacingRule.Exactly, // Use Exactly with LineSpacing = 24 for double
            "atLeast" => LineSpacingRule.AtLeast,
            "exactly" => LineSpacingRule.Exactly,
            "multiple" => LineSpacingRule.Multiple,
            _ => LineSpacingRule.Exactly
        };
    }
    
    private TabAlignment GetTabAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => TabAlignment.Left,
            "center" => TabAlignment.Center,
            "right" => TabAlignment.Right,
            "decimal" => TabAlignment.Decimal,
            "bar" => TabAlignment.Bar,
            "clear" => TabAlignment.Clear,
            _ => TabAlignment.Left
        };
    }
    
    private TabLeader GetTabLeader(string leader)
    {
        return leader.ToLower() switch
        {
            "none" => TabLeader.None,
            "dots" => TabLeader.Dots,
            "dashes" => TabLeader.Dashes,
            "line" => TabLeader.Line,
            "heavy" => TabLeader.Heavy,
            "middleDot" => TabLeader.MiddleDot,
            _ => TabLeader.None
        };
    }
    
    private System.Drawing.Color ParseColor(string colorStr)
    {
        if (string.IsNullOrEmpty(colorStr))
            return System.Drawing.Color.Black;
        
        // Remove # if present
        colorStr = colorStr.TrimStart('#');
        
        if (colorStr.Length == 6)
        {
            // RGB hex format
            var r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
            var g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
            var b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        
        return System.Drawing.Color.Black;
    }
}

