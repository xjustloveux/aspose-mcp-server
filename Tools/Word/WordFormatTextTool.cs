using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordFormatTextTool : IAsposeTool
{
    public string Description => "Format text at Run level (more granular than paragraph formatting)";

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
            runIndex = new
            {
                type = "number",
                description = "Run index within the paragraph (0-based). If not provided, formats all runs in the paragraph."
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional). If fontNameAscii and fontNameFarEast are provided, this will be used as fallback."
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
                description = "Font size in points (optional)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold (optional)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic (optional)"
            },
            underline = new
            {
                type = "string",
                description = "Underline style: none, single, double, dotted, dash (optional)",
                @enum = new[] { "none", "single", "double", "dotted", "dash" }
            },
            color = new
            {
                type = "string",
                description = "Text color (hex format like #FF0000 or color name like 'red', optional)"
            },
            strikethrough = new
            {
                type = "boolean",
                description = "Strikethrough (optional)"
            },
            superscript = new
            {
                type = "boolean",
                description = "Superscript (optional)"
            },
            subscript = new
            {
                type = "boolean",
                description = "Subscript (optional)"
            }
        },
        required = new[] { "path", "paragraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var runIndex = arguments?["runIndex"]?.GetValue<int?>();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var underline = arguments?["underline"]?.GetValue<string>();
        var color = arguments?["color"]?.GetValue<string>();
        var strikethrough = arguments?["strikethrough"]?.GetValue<bool?>();
        var superscript = arguments?["superscript"]?.GetValue<bool?>();
        var subscript = arguments?["subscript"]?.GetValue<bool?>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var para = paragraphs[paragraphIndex] as Paragraph;
        if (para == null)
        {
            throw new InvalidOperationException($"無法找到索引 {paragraphIndex} 的段落");
        }
        
        var runs = para.GetChildNodes(NodeType.Run, false);
        if (runs == null || runs.Count == 0)
        {
            throw new InvalidOperationException($"段落 #{paragraphIndex} 中沒有 Run 節點");
        }
        
        var changes = new List<string>();
        var runsToFormat = new List<Run>();
        
        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
            {
                throw new ArgumentException($"Run 索引 {runIndex.Value} 超出範圍 (段落共有 {runs.Count} 個 Run)");
            }
            var run = runs[runIndex.Value] as Run;
            if (run != null)
            {
                runsToFormat.Add(run);
            }
        }
        else
        {
            // Format all runs in the paragraph
            foreach (Node node in runs)
            {
                if (node is Run run)
                {
                    runsToFormat.Add(run);
                }
            }
        }
        
        foreach (var run in runsToFormat)
        {
            // Set font names (priority: fontNameAscii/fontNameFarEast > fontName)
            if (!string.IsNullOrEmpty(fontNameAscii))
            {
                run.Font.NameAscii = fontNameAscii;
                changes.Add($"字型（英文）: {fontNameAscii}");
            }
            
            if (!string.IsNullOrEmpty(fontNameFarEast))
            {
                run.Font.NameFarEast = fontNameFarEast;
                changes.Add($"字型（中文）: {fontNameFarEast}");
            }
            
            if (!string.IsNullOrEmpty(fontName))
            {
                // If fontNameAscii/FarEast are not set, use fontName for both
                if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                {
                    run.Font.Name = fontName;
                    changes.Add($"字型: {fontName}");
                }
                else
                {
                    // If only one is set, use fontName as fallback for the other
                    if (string.IsNullOrEmpty(fontNameAscii))
                    {
                        run.Font.NameAscii = fontName;
                    }
                    if (string.IsNullOrEmpty(fontNameFarEast))
                    {
                        run.Font.NameFarEast = fontName;
                    }
                }
            }
            
            if (fontSize.HasValue)
            {
                run.Font.Size = fontSize.Value;
                changes.Add($"字型大小: {fontSize.Value} 點");
            }
            
            if (bold.HasValue)
            {
                run.Font.Bold = bold.Value;
                changes.Add($"粗體: {(bold.Value ? "是" : "否")}");
            }
            
            if (italic.HasValue)
            {
                run.Font.Italic = italic.Value;
                changes.Add($"斜體: {(italic.Value ? "是" : "否")}");
            }
            
            if (!string.IsNullOrEmpty(underline))
            {
                run.Font.Underline = underline.ToLower() switch
                {
                    "single" => Underline.Single,
                    "double" => Underline.Double,
                    "dotted" => Underline.Dotted,
                    "dash" => Underline.Dash,
                    "none" => Underline.None,
                    _ => Underline.None
                };
                changes.Add($"底線: {underline}");
            }
            
            if (!string.IsNullOrEmpty(color))
            {
                try
                {
                    if (color.StartsWith("#"))
                        color = color.Substring(1);
                    
                    if (color.Length == 6)
                    {
                        int r = Convert.ToInt32(color.Substring(0, 2), 16);
                        int g = Convert.ToInt32(color.Substring(2, 2), 16);
                        int b = Convert.ToInt32(color.Substring(4, 2), 16);
                        run.Font.Color = System.Drawing.Color.FromArgb(r, g, b);
                        changes.Add($"顏色: #{color}");
                    }
                    else
                    {
                        run.Font.Color = System.Drawing.Color.FromName(color);
                        changes.Add($"顏色: {color}");
                    }
                }
                catch
                {
                    // Ignore color parsing errors
                }
            }
            
            if (strikethrough.HasValue)
            {
                run.Font.StrikeThrough = strikethrough.Value;
                changes.Add($"刪除線: {(strikethrough.Value ? "是" : "否")}");
            }
            
            if (superscript.HasValue && superscript.Value)
            {
                run.Font.Position = 6; // Superscript
                changes.Add("上標: 是");
            }
            
            if (subscript.HasValue && subscript.Value)
            {
                run.Font.Position = -6; // Subscript
                changes.Add("下標: 是");
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"成功設定 Run 層級格式\n";
        result += $"段落索引: {paragraphIndex}\n";
        if (runIndex.HasValue)
        {
            result += $"Run 索引: {runIndex.Value}\n";
        }
        else
        {
            result += $"格式化的 Run 數: {runsToFormat.Count}\n";
        }
        if (changes.Count > 0)
        {
            result += $"變更內容: {string.Join("、", changes.Distinct())}\n";
        }
        else
        {
            result += "未提供變更參數\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

