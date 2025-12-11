using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for paragraph operations in Word documents
/// Merges: WordInsertParagraphTool, WordDeleteParagraphTool, WordEditParagraphTool,
/// WordGetParagraphsTool, WordGetParagraphFormatTool, WordCopyParagraphFormatTool, WordMergeParagraphsTool
/// </summary>
public class WordParagraphTool : IAsposeTool
{
    public string Description => "Manage paragraphs in Word documents: insert, delete, edit format, get info, copy format, merge";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'insert', 'delete', 'edit', 'get', 'get_format', 'copy_format', 'merge'",
                @enum = new[] { "insert", "delete", "edit", "get", "get_format", "copy_format", "merge" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for write operations)"
            },
            // Common parameters
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based, required for delete, edit, get_format operations, optional for insert/get operations)"
            },
            // Insert parameters
            text = new
            {
                type = "string",
                description = "Text content for the paragraph (required for insert operation)"
            },
            styleName = new
            {
                type = "string",
                description = "Style name to apply (e.g., 'Heading 1', '標題1', 'Normal', optional, for insert/edit operations)"
            },
            alignment = new
            {
                type = "string",
                description = "Text alignment: left, center, right, justify (optional, for insert/edit operations)",
                @enum = new[] { "left", "center", "right", "justify" }
            },
            // Get parameters
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, for get operation)"
            },
            includeEmpty = new
            {
                type = "boolean",
                description = "Include empty paragraphs (optional, default: true, for get operation)"
            },
            styleFilter = new
            {
                type = "string",
                description = "Filter by style name (optional, for get operation)"
            },
            // Get format parameters
            includeRunDetails = new
            {
                type = "boolean",
                description = "Include detailed run-level formatting (optional, default: true, for get_format operation)"
            },
            // Edit parameters
            fontName = new
            {
                type = "string",
                description = "Font name (e.g., '標楷體', 'Arial', optional, for edit operation)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, optional, for edit operation)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean, optional, for edit operation)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (optional, for edit operation)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (optional, for edit operation)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text (optional, for edit operation)"
            },
            underline = new
            {
                type = "boolean",
                description = "Underline text (optional, for edit operation)"
            },
            color = new
            {
                type = "string",
                description = "Font color (hex format, e.g., '000000' for black, optional, for edit operation)"
            },
            indentLeft = new
            {
                type = "number",
                description = "Left indent in points (optional, for edit operation)"
            },
            indentRight = new
            {
                type = "number",
                description = "Right indent in points (optional, for edit operation)"
            },
            firstLineIndent = new
            {
                type = "number",
                description = "First line indent in points (positive for indent, negative for hanging, optional, for edit operation)"
            },
            spaceBefore = new
            {
                type = "number",
                description = "Space before paragraph in points (optional, for edit operation)"
            },
            spaceAfter = new
            {
                type = "number",
                description = "Space after paragraph in points (optional, for edit operation)"
            },
            lineSpacing = new
            {
                type = "number",
                description = "Line spacing (points or multiplier depending on lineSpacingRule, optional, for edit operation)"
            },
            lineSpacingRule = new
            {
                type = "string",
                description = "Line spacing rule: single, oneAndHalf, double, atLeast, exactly, multiple (optional, for edit operation)",
                @enum = new[] { "single", "oneAndHalf", "double", "atLeast", "exactly", "multiple" }
            },
            tabStops = new
            {
                type = "array",
                description = "Custom tab stops (optional, for edit operation)",
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
            },
            // Copy format parameters
            sourceParagraphIndex = new
            {
                type = "number",
                description = "Source paragraph index (0-based, required for copy_format operation)"
            },
            targetParagraphIndex = new
            {
                type = "number",
                description = "Target paragraph index (0-based, required for copy_format operation)"
            },
            // Merge parameters
            startParagraphIndex = new
            {
                type = "number",
                description = "Start paragraph index (0-based, inclusive, required for merge operation)"
            },
            endParagraphIndex = new
            {
                type = "number",
                description = "End paragraph index (0-based, inclusive, required for merge operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        return operation switch
        {
            "insert" => await InsertParagraphAsync(arguments, path, outputPath),
            "delete" => await DeleteParagraphAsync(arguments, path, outputPath),
            "edit" => await EditParagraphAsync(arguments, path, outputPath),
            "get" => await GetParagraphsAsync(arguments, path),
            "get_format" => await GetParagraphFormatAsync(arguments, path),
            "copy_format" => await CopyParagraphFormatAsync(arguments, path, outputPath),
            "merge" => await MergeParagraphsAsync(arguments, path, outputPath),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> InsertParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var styleName = arguments?["styleName"]?.GetValue<string>();
        var alignment = arguments?["alignment"]?.GetValue<string>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        Paragraph? targetPara = null;
        string insertPosition = "文檔末尾";
        
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                {
                    targetPara = paragraphs[0] as Paragraph;
                    insertPosition = "文檔開頭";
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                insertPosition = $"段落 #{paragraphIndex.Value} 之後";
            }
            else
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
        }

        var para = new Paragraph(doc);
        var run = new Run(doc, text);
        para.AppendChild(run);

        if (!string.IsNullOrEmpty(styleName))
        {
            try
            {
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
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法應用樣式 '{styleName}': {ex.Message}，可用樣式請使用 word_get_styles 工具查看", ex);
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

        if (targetPara != null)
        {
            if (paragraphIndex!.Value == -1)
            {
                targetPara.ParentNode.InsertBefore(para, targetPara);
            }
            else
            {
                targetPara.ParentNode.InsertAfter(para, targetPara);
            }
        }
        else
        {
            var body = doc.FirstSection.Body;
            body.AppendChild(para);
        }

        doc.Save(outputPath);

        var result = $"成功插入段落\n";
        result += $"插入位置: {insertPosition}\n";
        if (!string.IsNullOrEmpty(styleName))
        {
            result += $"應用樣式: {styleName}\n";
        }
        if (!string.IsNullOrEmpty(alignment))
        {
            result += $"對齊方式: {alignment}\n";
        }
        result += $"文檔段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> DeleteParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }

        var paragraphToDelete = paragraphs[paragraphIndex] as Paragraph;
        if (paragraphToDelete == null)
        {
            throw new InvalidOperationException($"無法獲取索引 {paragraphIndex} 的段落");
        }

        var textPreview = paragraphToDelete.GetText().Trim();
        if (textPreview.Length > 50)
        {
            textPreview = textPreview.Substring(0, 50) + "...";
        }
        
        paragraphToDelete.Remove();

        doc.Save(outputPath);

        var result = $"成功刪除段落 #{paragraphIndex}\n";
        if (!string.IsNullOrEmpty(textPreview))
        {
            result += $"內容預覽: {textPreview}\n";
        }
        result += $"文檔剩餘段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> EditParagraphAsync(JsonObject? arguments, string path, string outputPath)
    {
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
                paraFormat.LineSpacing = 12;
            }
            else if (lineSpacingRule == "oneAndHalf")
            {
                paraFormat.LineSpacing = 18;
            }
            else if (lineSpacingRule == "double")
            {
                paraFormat.LineSpacing = 24;
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

    private async Task<string> GetParagraphsAsync(JsonObject? arguments, string path)
    {
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var includeEmpty = arguments?["includeEmpty"]?.GetValue<bool?>() ?? true;
        var styleFilter = arguments?["styleFilter"]?.GetValue<string>();

        var doc = new Document(path);
        var sb = new StringBuilder();

        List<Paragraph> paragraphs;
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }
            paragraphs = doc.Sections[sectionIndex.Value].Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        }
        else
        {
            paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        }

        if (!includeEmpty)
        {
            paragraphs = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.GetText())).ToList();
        }

        if (!string.IsNullOrEmpty(styleFilter))
        {
            paragraphs = paragraphs.Where(p => p.ParagraphFormat.Style?.Name == styleFilter).ToList();
        }

        sb.AppendLine($"=== Paragraphs ({paragraphs.Count}) ===");
        sb.AppendLine();

        for (int i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var text = para.GetText().Trim();
            sb.AppendLine($"[{i}] Style: {para.ParagraphFormat.Style?.Name ?? "(none)"}");
            sb.AppendLine($"    Text: {text.Substring(0, Math.Min(100, text.Length))}{(text.Length > 100 ? "..." : "")}");
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> GetParagraphFormatAsync(JsonObject? arguments, string path)
    {
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var includeRunDetails = arguments?["includeRunDetails"]?.GetValue<bool>() ?? true;

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

        var result = new StringBuilder();
        result.AppendLine($"=== 段落 #{paragraphIndex} 格式資訊 ===\n");

        result.AppendLine("【基本資訊】");
        result.AppendLine($"段落文字: {para.GetText().Trim()}");
        result.AppendLine($"文字長度: {para.GetText().Trim().Length} 字元");
        result.AppendLine($"Run 數量: {para.Runs.Count}");
        result.AppendLine();

        var format = para.ParagraphFormat;
        result.AppendLine("【段落格式】");
        result.AppendLine($"樣式名稱: {format.StyleName}");
        result.AppendLine($"對齊方式: {format.Alignment}");
        result.AppendLine($"左縮排: {format.LeftIndent:F2} pt ({format.LeftIndent / 28.35:F2} cm)");
        result.AppendLine($"右縮排: {format.RightIndent:F2} pt ({format.RightIndent / 28.35:F2} cm)");
        result.AppendLine($"首行縮排: {format.FirstLineIndent:F2} pt ({format.FirstLineIndent / 28.35:F2} cm)");
        result.AppendLine($"段前間距: {format.SpaceBefore:F2} pt");
        result.AppendLine($"段後間距: {format.SpaceAfter:F2} pt");
        result.AppendLine($"行距: {format.LineSpacing:F2} pt");
        result.AppendLine($"行距規則: {format.LineSpacingRule}");
        result.AppendLine();

        if (para.ListFormat != null && para.ListFormat.IsListItem)
        {
            result.AppendLine("【列表格式】");
            result.AppendLine($"是列表項: 是");
            result.AppendLine($"列表層級: {para.ListFormat.ListLevelNumber}");
            if (para.ListFormat.List != null)
            {
                result.AppendLine($"列表 ID: {para.ListFormat.List.ListId}");
            }
            result.AppendLine();
        }

        if (format.Borders.Count > 0)
        {
            result.AppendLine("【邊框】");
            if (format.Borders.Top.LineStyle != LineStyle.None)
                result.AppendLine($"上邊框: {format.Borders.Top.LineStyle}, {format.Borders.Top.LineWidth} pt, 顏色: {format.Borders.Top.Color.Name}");
            if (format.Borders.Bottom.LineStyle != LineStyle.None)
                result.AppendLine($"下邊框: {format.Borders.Bottom.LineStyle}, {format.Borders.Bottom.LineWidth} pt, 顏色: {format.Borders.Bottom.Color.Name}");
            if (format.Borders.Left.LineStyle != LineStyle.None)
                result.AppendLine($"左邊框: {format.Borders.Left.LineStyle}, {format.Borders.Left.LineWidth} pt, 顏色: {format.Borders.Left.Color.Name}");
            if (format.Borders.Right.LineStyle != LineStyle.None)
                result.AppendLine($"右邊框: {format.Borders.Right.LineStyle}, {format.Borders.Right.LineWidth} pt, 顏色: {format.Borders.Right.Color.Name}");
            result.AppendLine();
        }

        if (format.Shading.BackgroundPatternColor.ToArgb() != System.Drawing.Color.Empty.ToArgb())
        {
            result.AppendLine("【背景色】");
            var color = format.Shading.BackgroundPatternColor;
            result.AppendLine($"背景色: #{color.R:X2}{color.G:X2}{color.B:X2}");
            result.AppendLine();
        }

        if (format.TabStops.Count > 0)
        {
            result.AppendLine("【Tab 停駐點】");
            for (int i = 0; i < format.TabStops.Count; i++)
            {
                var tab = format.TabStops[i];
                result.AppendLine($"  Tab {i + 1}: 位置={tab.Position:F2} pt, 對齊={tab.Alignment}, 前導字元={tab.Leader}");
            }
            result.AppendLine();
        }

        if (para.Runs.Count > 0)
        {
            var firstRun = para.Runs[0];
            result.AppendLine("【字型格式（第一個 Run）】");
            
            if (firstRun.Font.NameAscii != firstRun.Font.NameFarEast)
            {
                result.AppendLine($"字體（英文）: {firstRun.Font.NameAscii}");
                result.AppendLine($"字體（中文）: {firstRun.Font.NameFarEast}");
            }
            else
            {
                result.AppendLine($"字體: {firstRun.Font.Name}");
            }
            
            result.AppendLine($"字號: {firstRun.Font.Size} pt");
            
            if (firstRun.Font.Bold) result.AppendLine("粗體: 是");
            if (firstRun.Font.Italic) result.AppendLine("斜體: 是");
            if (firstRun.Font.Underline != Underline.None) result.AppendLine($"底線: {firstRun.Font.Underline}");
            if (firstRun.Font.StrikeThrough) result.AppendLine("刪除線: 是");
            if (firstRun.Font.Superscript) result.AppendLine("上標: 是");
            if (firstRun.Font.Subscript) result.AppendLine("下標: 是");
            
            if (firstRun.Font.Color.ToArgb() != System.Drawing.Color.Empty.ToArgb())
            {
                var color = firstRun.Font.Color;
                result.AppendLine($"顏色: #{color.R:X2}{color.G:X2}{color.B:X2}");
            }
            
            if (firstRun.Font.HighlightColor != System.Drawing.Color.Empty)
            {
                result.AppendLine($"螢光筆: {firstRun.Font.HighlightColor.Name}");
            }
            result.AppendLine();
        }

        if (includeRunDetails && para.Runs.Count > 1)
        {
            result.AppendLine("【Run 詳細資訊】");
            result.AppendLine($"共 {para.Runs.Count} 個 Run:");
            
            for (int i = 0; i < Math.Min(para.Runs.Count, 10); i++)
            {
                var run = para.Runs[i];
                result.AppendLine($"\n  Run #{i}:");
                result.AppendLine($"    文字: {run.Text.Replace("\r", "\\r").Replace("\n", "\\n")}");
                
                if (run.Font.NameAscii != run.Font.NameFarEast)
                {
                    result.AppendLine($"    字體（英文）: {run.Font.NameAscii}");
                    result.AppendLine($"    字體（中文）: {run.Font.NameFarEast}");
                }
                else
                {
                    result.AppendLine($"    字體: {run.Font.Name}");
                }
                
                result.AppendLine($"    字號: {run.Font.Size} pt");
                
                var styles = new List<string>();
                if (run.Font.Bold) styles.Add("粗體");
                if (run.Font.Italic) styles.Add("斜體");
                if (run.Font.Underline != Underline.None) styles.Add($"底線({run.Font.Underline})");
                if (styles.Count > 0)
                    result.AppendLine($"    樣式: {string.Join(", ", styles)}");
            }
            
            if (para.Runs.Count > 10)
            {
                result.AppendLine($"\n  ... 還有 {para.Runs.Count - 10} 個 Run（已省略）");
            }
            result.AppendLine();
        }

        result.AppendLine("【JSON 格式（可用於 word_edit_paragraph）】");
        result.AppendLine("{");
        result.AppendLine($"  \"alignment\": \"{format.Alignment.ToString().ToLower()}\",");
        result.AppendLine($"  \"leftIndent\": {format.LeftIndent:F2},");
        result.AppendLine($"  \"rightIndent\": {format.RightIndent:F2},");
        result.AppendLine($"  \"firstLineIndent\": {format.FirstLineIndent:F2},");
        result.AppendLine($"  \"spaceBefore\": {format.SpaceBefore:F2},");
        result.AppendLine($"  \"spaceAfter\": {format.SpaceAfter:F2},");
        result.AppendLine($"  \"lineSpacing\": {format.LineSpacing:F2}");
        
        if (para.Runs.Count > 0)
        {
            var run = para.Runs[0];
            result.AppendLine($"  \"fontNameAscii\": \"{run.Font.NameAscii}\",");
            result.AppendLine($"  \"fontNameFarEast\": \"{run.Font.NameFarEast}\",");
            result.AppendLine($"  \"fontSize\": {run.Font.Size},");
            result.AppendLine($"  \"bold\": {run.Font.Bold.ToString().ToLower()},");
            result.AppendLine($"  \"italic\": {run.Font.Italic.ToString().ToLower()}");
        }
        
        result.AppendLine("}");

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> CopyParagraphFormatAsync(JsonObject? arguments, string path, string outputPath)
    {
        var sourceParagraphIndex = arguments?["sourceParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("sourceParagraphIndex is required");
        var targetParagraphIndex = arguments?["targetParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("targetParagraphIndex is required");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (sourceParagraphIndex < 0 || sourceParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"來源段落索引 {sourceParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (targetParagraphIndex < 0 || targetParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"目標段落索引 {targetParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var sourcePara = paragraphs[sourceParagraphIndex] as Paragraph;
        var targetPara = paragraphs[targetParagraphIndex] as Paragraph;
        
        if (sourcePara == null || targetPara == null)
        {
            throw new InvalidOperationException("無法獲取段落");
        }
        
        targetPara.ParagraphFormat.StyleName = sourcePara.ParagraphFormat.StyleName;
        targetPara.ParagraphFormat.Alignment = sourcePara.ParagraphFormat.Alignment;
        targetPara.ParagraphFormat.LeftIndent = sourcePara.ParagraphFormat.LeftIndent;
        targetPara.ParagraphFormat.RightIndent = sourcePara.ParagraphFormat.RightIndent;
        targetPara.ParagraphFormat.FirstLineIndent = sourcePara.ParagraphFormat.FirstLineIndent;
        targetPara.ParagraphFormat.SpaceBefore = sourcePara.ParagraphFormat.SpaceBefore;
        targetPara.ParagraphFormat.SpaceAfter = sourcePara.ParagraphFormat.SpaceAfter;
        targetPara.ParagraphFormat.LineSpacing = sourcePara.ParagraphFormat.LineSpacing;
        targetPara.ParagraphFormat.LineSpacingRule = sourcePara.ParagraphFormat.LineSpacingRule;
        
        targetPara.ParagraphFormat.TabStops.Clear();
        for (int i = 0; i < sourcePara.ParagraphFormat.TabStops.Count; i++)
        {
            var tabStop = sourcePara.ParagraphFormat.TabStops[i];
            targetPara.ParagraphFormat.TabStops.Add(tabStop.Position, tabStop.Alignment, tabStop.Leader);
        }
        
        doc.Save(outputPath);
        
        var result = $"成功複製段落格式\n";
        result += $"來源段落: #{sourceParagraphIndex}\n";
        result += $"目標段落: #{targetParagraphIndex}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> MergeParagraphsAsync(JsonObject? arguments, string path, string outputPath)
    {
        var startParagraphIndex = arguments?["startParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("startParagraphIndex is required");
        var endParagraphIndex = arguments?["endParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("endParagraphIndex is required");

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"起始段落索引 {startParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"結束段落索引 {endParagraphIndex} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        if (startParagraphIndex > endParagraphIndex)
        {
            throw new ArgumentException($"起始段落索引 {startParagraphIndex} 不能大於結束段落索引 {endParagraphIndex}");
        }
        
        if (startParagraphIndex == endParagraphIndex)
        {
            throw new ArgumentException("起始和結束段落索引相同，無需合併");
        }
        
        var startPara = paragraphs[startParagraphIndex] as Paragraph;
        if (startPara == null)
        {
            throw new InvalidOperationException("無法獲取起始段落");
        }
        
        for (int i = startParagraphIndex + 1; i <= endParagraphIndex; i++)
        {
            var para = paragraphs[i] as Paragraph;
            if (para != null)
            {
                if (startPara.Runs.Count > 0)
                {
                    var spaceRun = new Run(doc, " ");
                    startPara.AppendChild(spaceRun);
                }
                
                var runsToMove = para.Runs.ToArray();
                foreach (var run in runsToMove)
                {
                    startPara.AppendChild(run);
                }
                
                para.Remove();
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"成功合併段落\n";
        result += $"合併範圍: 段落 #{startParagraphIndex} 到 #{endParagraphIndex}\n";
        result += $"合併段落數: {endParagraphIndex - startParagraphIndex + 1}\n";
        result += $"文檔剩餘段落數: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
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
            "single" => LineSpacingRule.Exactly,
            "oneAndHalf" => LineSpacingRule.Exactly,
            "double" => LineSpacingRule.Exactly,
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
        
        colorStr = colorStr.TrimStart('#');
        
        if (colorStr.Length == 6)
        {
            var r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
            var g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
            var b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        
        return System.Drawing.Color.Black;
    }
}

