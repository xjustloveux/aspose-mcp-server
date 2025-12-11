using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for text operations in Word documents
/// Merges: WordAddTextTool, WordDeleteTextTool, WordReplaceTextTool, WordSearchTextTool, 
/// WordFormatTextTool, WordInsertTextAtPositionTool, WordDeleteTextRangeTool, WordAddTextWithStyleTool
/// </summary>
public class WordTextTool : IAsposeTool
{
    public string Description => "Perform text operations in Word documents: add, delete, replace, search, format, insert at position";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'delete', 'replace', 'search', 'format', 'insert_at_position', 'delete_range', 'add_with_style'",
                @enum = new[] { "add", "delete", "replace", "search", "format", "insert_at_position", "delete_range", "add_with_style" }
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
            text = new
            {
                type = "string",
                description = "Text content (required for add, replace, insert_at_position, add_with_style operations)"
            },
            // Add/AddWithStyle parameters
            fontName = new
            {
                type = "string",
                description = "Font name (optional, e.g., 'Arial', '微軟雅黑')"
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
                description = "Font size (optional)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (optional)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text (optional)"
            },
            underline = new
            {
                type = "string",
                description = "Underline style: none, single, double, dotted, dash (optional, for format operation)",
                @enum = new[] { "none", "single", "double", "dotted", "dash" }
            },
            color = new
            {
                type = "string",
                description = "Text color (hex format like 'FF0000' or '#FF0000' for red, or name like 'Red', 'Blue', optional)"
            },
            strikethrough = new
            {
                type = "boolean",
                description = "Strikethrough (optional, for format operation)"
            },
            superscript = new
            {
                type = "boolean",
                description = "Superscript (optional, for format operation)"
            },
            subscript = new
            {
                type = "boolean",
                description = "Subscript (optional, for format operation)"
            },
            // Replace parameters
            find = new
            {
                type = "string",
                description = "Text to find (required for replace operation)"
            },
            replace = new
            {
                type = "string",
                description = "Replacement text (required for replace operation)"
            },
            useRegex = new
            {
                type = "boolean",
                description = "Use regex matching (optional, for replace/search operations)"
            },
            // Search parameters
            searchText = new
            {
                type = "string",
                description = "Text to search for (required for search operation)"
            },
            caseSensitive = new
            {
                type = "boolean",
                description = "Case sensitive search (optional, default: false, for search operation)"
            },
            maxResults = new
            {
                type = "number",
                description = "Maximum number of results to return (optional, default: 50, for search operation)"
            },
            contextLength = new
            {
                type = "number",
                description = "Number of characters to show before and after match for context (optional, default: 50, for search operation)"
            },
            // Delete parameters
            startParagraphIndex = new
            {
                type = "number",
                description = "Start paragraph index (0-based, required for delete operation)"
            },
            startRunIndex = new
            {
                type = "number",
                description = "Start run index within start paragraph (0-based, optional, default: 0, for delete operation)"
            },
            endParagraphIndex = new
            {
                type = "number",
                description = "End paragraph index (0-based, inclusive, required for delete operation)"
            },
            endRunIndex = new
            {
                type = "number",
                description = "End run index within end paragraph (0-based, inclusive, optional, default: last run, for delete operation)"
            },
            // Format parameters
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based, required for format operation)"
            },
            runIndex = new
            {
                type = "number",
                description = "Run index within the paragraph (0-based, optional, formats all runs if not provided, for format operation)"
            },
            // Insert at position parameters
            insertParagraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based, required for insert_at_position operation)"
            },
            charIndex = new
            {
                type = "number",
                description = "Character index within paragraph (0-based, required for insert_at_position operation)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0, for insert_at_position/delete_range operations)"
            },
            insertBefore = new
            {
                type = "boolean",
                description = "Insert before position (optional, default: false, inserts after, for insert_at_position operation)"
            },
            // Delete range parameters
            startCharIndex = new
            {
                type = "number",
                description = "Start character index within paragraph (0-based, required for delete_range operation)"
            },
            endCharIndex = new
            {
                type = "number",
                description = "End character index within paragraph (0-based, required for delete_range operation)"
            },
            // AddWithStyle parameters
            styleName = new
            {
                type = "string",
                description = "Style name to apply (e.g., 'Heading 1', '標題1', 'Normal', optional, for add_with_style operation)"
            },
            alignment = new
            {
                type = "string",
                description = "Text alignment: left, center, right, justify (optional, for add_with_style operation)",
                @enum = new[] { "left", "center", "right", "justify" }
            },
            indentLevel = new
            {
                type = "number",
                description = "Indentation level (0-8, where each level = 36 points / 0.5 inch, optional, for add_with_style operation)"
            },
            leftIndent = new
            {
                type = "number",
                description = "Left indentation in points (optional, for add_with_style operation)"
            },
            firstLineIndent = new
            {
                type = "number",
                description = "First line indentation in points (positive = indent first line, negative = hanging indent, optional, for add_with_style operation)"
            },
            paragraphIndexForAdd = new
            {
                type = "number",
                description = "Index of the paragraph to insert after (0-based, optional, for add_with_style operation). Use -1 to insert at the beginning."
            },
            tabStops = new
            {
                type = "array",
                description = "Custom tab stops (optional, for add_with_style operation)",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        position = new { type = "number", description = "Tab stop position in points" },
                        alignment = new { type = "string", description = "Tab alignment: Left, Center, Right, Decimal, Bar", @enum = new[] { "Left", "Center", "Right", "Decimal", "Bar" } },
                        leader = new { type = "string", description = "Tab leader: None, Dots, Dashes, Line, Heavy, MiddleDot", @enum = new[] { "None", "Dots", "Dashes", "Line", "Heavy", "MiddleDot" } }
                    }
                }
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
            "add" => await AddTextAsync(arguments, path, outputPath),
            "delete" => await DeleteTextAsync(arguments, path, outputPath),
            "replace" => await ReplaceTextAsync(arguments, path, outputPath),
            "search" => await SearchTextAsync(arguments, path),
            "format" => await FormatTextAsync(arguments, path, outputPath),
            "insert_at_position" => await InsertTextAtPositionAsync(arguments, path, outputPath),
            "delete_range" => await DeleteTextRangeAsync(arguments, path, outputPath),
            "add_with_style" => await AddTextWithStyleAsync(arguments, path, outputPath),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double>();
        var bold = arguments?["bold"]?.GetValue<bool>() ?? false;
        var italic = arguments?["italic"]?.GetValue<bool>() ?? false;
        var color = arguments?["color"]?.GetValue<string>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(fontName))
            builder.Font.Name = fontName;
        if (fontSize.HasValue)
            builder.Font.Size = fontSize.Value;
        builder.Font.Bold = bold;
        builder.Font.Italic = italic;
        
        if (!string.IsNullOrEmpty(color))
        {
            try
            {
                if (color.StartsWith("#"))
                    color = color.TrimStart('#');
                
                if (color.Length == 6)
                {
                    var r = Convert.ToInt32(color.Substring(0, 2), 16);
                    var g = Convert.ToInt32(color.Substring(2, 2), 16);
                    var b = Convert.ToInt32(color.Substring(4, 2), 16);
                    builder.Font.Color = System.Drawing.Color.FromArgb(r, g, b);
                }
                else
                {
                    builder.Font.Color = System.Drawing.Color.FromName(color);
                }
            }
            catch
            {
                // Ignore invalid color
            }
        }

        builder.Writeln(text);
        doc.Save(outputPath);

        return await Task.FromResult($"Text added to document: {outputPath}");
    }

    private async Task<string> DeleteTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        var startParagraphIndex = arguments?["startParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("startParagraphIndex is required");
        var startRunIndex = arguments?["startRunIndex"]?.GetValue<int>() ?? 0;
        var endParagraphIndex = arguments?["endParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("endParagraphIndex is required");
        var endRunIndex = arguments?["endRunIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count ||
            endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count ||
            startParagraphIndex > endParagraphIndex)
        {
            throw new ArgumentException($"段落索引超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var startPara = paragraphs[startParagraphIndex] as Paragraph;
        var endPara = paragraphs[endParagraphIndex] as Paragraph;
        
        if (startPara == null || endPara == null)
        {
            throw new InvalidOperationException("無法找到指定的段落");
        }
        
        // Get deleted text preview before deletion
        string deletedText = "";
        try
        {
            var startRuns = startPara.GetChildNodes(NodeType.Run, false);
            var endRuns = endPara.GetChildNodes(NodeType.Run, false);
            
            if (startParagraphIndex == endParagraphIndex)
            {
                if (startRuns != null && startRuns.Count > 0)
                {
                    var actualEndRunIndex = endRunIndex ?? (startRuns.Count - 1);
                    if (startRunIndex >= 0 && startRunIndex < startRuns.Count &&
                        actualEndRunIndex >= 0 && actualEndRunIndex < startRuns.Count &&
                        startRunIndex <= actualEndRunIndex)
                    {
                        for (int i = startRunIndex; i <= actualEndRunIndex; i++)
                        {
                            if (startRuns[i] is Run run)
                            {
                                deletedText += run.Text;
                            }
                        }
                    }
                }
            }
            else
            {
                if (startRuns != null && startRuns.Count > startRunIndex)
                {
                    for (int i = startRunIndex; i < startRuns.Count; i++)
                    {
                        if (startRuns[i] is Run run)
                        {
                            deletedText += run.Text;
                        }
                    }
                }
                
                for (int p = startParagraphIndex + 1; p < endParagraphIndex; p++)
                {
                    var para = paragraphs[p] as Paragraph;
                    if (para != null)
                    {
                        deletedText += para.GetText();
                    }
                }
                
                if (endRuns != null && endRuns.Count > 0)
                {
                    var actualEndRunIndex = endRunIndex ?? (endRuns.Count - 1);
                    for (int i = 0; i <= actualEndRunIndex && i < endRuns.Count; i++)
                    {
                        if (endRuns[i] is Run run)
                        {
                            deletedText += run.Text;
                        }
                    }
                }
            }
        }
        catch
        {
            // Ignore preview errors
        }
        
        // Delete text
        if (startParagraphIndex == endParagraphIndex)
        {
            var runs = startPara.GetChildNodes(NodeType.Run, false);
            if (runs != null && runs.Count > 0)
            {
                var actualEndRunIndex = endRunIndex ?? (runs.Count - 1);
                if (startRunIndex >= 0 && startRunIndex < runs.Count &&
                    actualEndRunIndex >= 0 && actualEndRunIndex < runs.Count &&
                    startRunIndex <= actualEndRunIndex)
                {
                    for (int i = actualEndRunIndex; i >= startRunIndex; i--)
                    {
                        runs[i]?.Remove();
                    }
                }
            }
        }
        else
        {
            var startRuns = startPara.GetChildNodes(NodeType.Run, false);
            if (startRuns != null && startRuns.Count > startRunIndex)
            {
                for (int i = startRuns.Count - 1; i >= startRunIndex; i--)
                {
                    startRuns[i]?.Remove();
                }
            }
            
            for (int p = endParagraphIndex - 1; p > startParagraphIndex; p--)
            {
                paragraphs[p]?.Remove();
            }
            
            var endRuns = endPara.GetChildNodes(NodeType.Run, false);
            if (endRuns != null && endRuns.Count > 0)
            {
                var actualEndRunIndex = endRunIndex ?? (endRuns.Count - 1);
                for (int i = actualEndRunIndex; i >= 0; i--)
                {
                    if (i < endRuns.Count)
                    {
                        endRuns[i]?.Remove();
                    }
                }
            }
        }
        
        doc.Save(outputPath);
        
        string preview = deletedText.Length > 50 ? deletedText.Substring(0, 50) + "..." : deletedText;
        
        var result = $"成功刪除文字\n";
        result += $"範圍: 段落 {startParagraphIndex} Run {startRunIndex} 到 段落 {endParagraphIndex} Run {endRunIndex ?? -1}\n";
        if (!string.IsNullOrEmpty(preview))
        {
            result += $"刪除內容預覽: {preview}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> ReplaceTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        var find = arguments?["find"]?.GetValue<string>() ?? throw new ArgumentException("find is required");
        var replace = arguments?["replace"]?.GetValue<string>() ?? throw new ArgumentException("replace is required");
        var useRegex = arguments?["useRegex"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        
        var options = new FindReplaceOptions();
        if (useRegex)
        {
            doc.Range.Replace(new Regex(find), replace, options);
        }
        else
        {
            doc.Range.Replace(find, replace, options);
        }

        doc.Save(outputPath);

        return await Task.FromResult($"Text replaced in document: {outputPath}");
    }

    private async Task<string> SearchTextAsync(JsonObject? arguments, string path)
    {
        var searchText = arguments?["searchText"]?.GetValue<string>() ?? throw new ArgumentException("searchText is required");
        var useRegex = arguments?["useRegex"]?.GetValue<bool>() ?? false;
        var caseSensitive = arguments?["caseSensitive"]?.GetValue<bool>() ?? false;
        var maxResults = arguments?["maxResults"]?.GetValue<int>() ?? 50;
        var contextLength = arguments?["contextLength"]?.GetValue<int>() ?? 50;

        var doc = new Document(path);
        var result = new StringBuilder();
        var matches = new List<(string text, int paragraphIndex, string context)>();

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        for (int i = 0; i < paragraphs.Count && matches.Count < maxResults; i++)
        {
            var para = paragraphs[i] as Paragraph;
            if (para == null) continue;
            
            var paraText = para.GetText();
            
            if (useRegex)
            {
                var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
                var regex = new Regex(searchText, options);
                var regexMatches = regex.Matches(paraText);
                
                foreach (Match match in regexMatches)
                {
                    if (matches.Count >= maxResults) break;
                    
                    var context = GetContext(paraText, match.Index, match.Length, contextLength);
                    matches.Add((match.Value, i, context));
                }
            }
            else
            {
                var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                int index = 0;
                
                while ((index = paraText.IndexOf(searchText, index, comparison)) != -1)
                {
                    if (matches.Count >= maxResults) break;
                    
                    var context = GetContext(paraText, index, searchText.Length, contextLength);
                    matches.Add((searchText, i, context));
                    index += searchText.Length;
                }
            }
        }

        result.AppendLine($"=== 搜尋結果 ===");
        result.AppendLine($"搜尋文字: {searchText}");
        result.AppendLine($"使用正則表達式: {(useRegex ? "是" : "否")}");
        result.AppendLine($"區分大小寫: {(caseSensitive ? "是" : "否")}");
        result.AppendLine($"找到 {matches.Count} 個匹配項{(matches.Count >= maxResults ? $" (限制前 {maxResults} 個)" : "")}\n");

        if (matches.Count == 0)
        {
            result.AppendLine("未找到匹配的文字");
        }
        else
        {
            for (int i = 0; i < matches.Count; i++)
            {
                var match = matches[i];
                result.AppendLine($"匹配 #{i + 1}:");
                result.AppendLine($"  位置: 段落 #{match.paragraphIndex}");
                result.AppendLine($"  匹配文字: {match.text}");
                result.AppendLine($"  上下文: ...{match.context}...");
                result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }

    private string GetContext(string text, int matchIndex, int matchLength, int contextLength)
    {
        int start = Math.Max(0, matchIndex - contextLength);
        int end = Math.Min(text.Length, matchIndex + matchLength + contextLength);
        
        var context = text.Substring(start, end - start);
        
        context = context.Replace("\r", "").Replace("\n", " ").Trim();
        
        int highlightStart = matchIndex - start;
        int highlightEnd = highlightStart + matchLength;
        
        if (highlightStart >= 0 && highlightEnd <= context.Length)
        {
            context = context.Substring(0, highlightStart) + 
                     "【" + context.Substring(highlightStart, matchLength) + "】" + 
                     context.Substring(highlightEnd);
        }
        
        return context;
    }

    private async Task<string> FormatTextAsync(JsonObject? arguments, string path, string outputPath)
    {
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
                if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                {
                    run.Font.Name = fontName;
                    changes.Add($"字型: {fontName}");
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
                run.Font.Position = 6;
                changes.Add("上標: 是");
            }
            
            if (subscript.HasValue && subscript.Value)
            {
                run.Font.Position = -6;
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

    private async Task<string> InsertTextAtPositionAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = arguments?["insertParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("insertParagraphIndex is required");
        var charIndex = arguments?["charIndex"]?.GetValue<int>() ?? throw new ArgumentException("charIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var insertBefore = arguments?["insertBefore"]?.GetValue<bool?>() ?? false;

        var doc = new Document(path);
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var section = doc.Sections[sectionIdx];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");
        }

        var para = paragraphs[paragraphIndex];
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var totalChars = 0;
        int targetRunIndex = -1;
        int targetRunCharIndex = 0;

        for (int i = 0; i < runs.Count; i++)
        {
            var runLength = runs[i].Text.Length;
            if (totalChars + runLength >= charIndex)
            {
                targetRunIndex = i;
                targetRunCharIndex = charIndex - totalChars;
                break;
            }
            totalChars += runLength;
        }

        if (targetRunIndex == -1)
        {
            var builder = new DocumentBuilder(doc);
            builder.MoveTo(para);
            builder.Write(text);
        }
        else
        {
            var targetRun = runs[targetRunIndex];
            targetRun.Text = targetRun.Text.Insert(targetRunCharIndex, text);
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Text inserted at position: {outputPath}");
    }

    private async Task<string> DeleteTextRangeAsync(JsonObject? arguments, string path, string outputPath)
    {
        var startParagraphIndex = arguments?["startParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("startParagraphIndex is required");
        var startCharIndex = arguments?["startCharIndex"]?.GetValue<int>() ?? throw new ArgumentException("startCharIndex is required");
        var endParagraphIndex = arguments?["endParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("endParagraphIndex is required");
        var endCharIndex = arguments?["endCharIndex"]?.GetValue<int>() ?? throw new ArgumentException("endCharIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var section = doc.Sections[sectionIdx];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count ||
            endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException("Paragraph indices out of range");
        }

        var startPara = paragraphs[startParagraphIndex];
        var endPara = paragraphs[endParagraphIndex];

        if (startParagraphIndex == endParagraphIndex)
        {
            var runs = startPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var totalChars = 0;
            int startRunIndex = -1, endRunIndex = -1;
            int startRunCharIndex = 0, endRunCharIndex = 0;

            for (int i = 0; i < runs.Count; i++)
            {
                var runLength = runs[i].Text.Length;
                if (startRunIndex == -1 && totalChars + runLength > startCharIndex)
                {
                    startRunIndex = i;
                    startRunCharIndex = startCharIndex - totalChars;
                }
                if (totalChars + runLength > endCharIndex)
                {
                    endRunIndex = i;
                    endRunCharIndex = endCharIndex - totalChars;
                    break;
                }
                totalChars += runLength;
            }

            if (startRunIndex >= 0 && endRunIndex >= 0)
            {
                if (startRunIndex == endRunIndex)
                {
                    var run = runs[startRunIndex];
                    run.Text = run.Text.Remove(startRunCharIndex, endRunCharIndex - startRunCharIndex);
                }
                else
                {
                    var startRun = runs[startRunIndex];
                    startRun.Text = startRun.Text.Substring(0, startRunCharIndex);

                    for (int i = startRunIndex + 1; i < endRunIndex; i++)
                    {
                        runs[i].Remove();
                    }

                    if (endRunIndex < runs.Count)
                    {
                        var endRun = runs[endRunIndex];
                        endRun.Text = endRun.Text.Substring(endRunCharIndex);
                    }
                }
            }
        }
        else
        {
            var startParaRuns = startPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var startRun = startParaRuns.LastOrDefault();
            if (startRun != null && startRun.Text.Length > startCharIndex)
            {
                startRun.Text = startRun.Text.Substring(0, startCharIndex);
            }

            for (int i = startParagraphIndex + 1; i < endParagraphIndex; i++)
            {
                paragraphs[i].Remove();
            }

            var endParaRuns = endPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            if (endParaRuns.Count > 0 && endCharIndex < endParaRuns[0].Text.Length)
            {
                endParaRuns[0].Text = endParaRuns[0].Text.Substring(endCharIndex);
                for (int i = 1; i < endParaRuns.Count; i++)
                {
                    endParaRuns[i].Remove();
                }
            }
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Text range deleted: {outputPath}");
    }

    private async Task<string> AddTextWithStyleAsync(JsonObject? arguments, string path, string outputPath)
    {
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
        var paragraphIndex = arguments?["paragraphIndexForAdd"]?.GetValue<int?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        Paragraph? targetPara = null;
        
        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
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
            builder.MoveToDocumentEnd();
        }

        var para = new Paragraph(doc);
        var run = new Run(doc, text);

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
            catch (Exception ex)
            {
                throw new InvalidOperationException($"無法應用樣式 '{styleName}': {ex.Message}，可用樣式請使用 word_get_styles 工具查看", ex);
            }
        }
        
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
        
        if (underline.HasValue)
            run.Font.Underline = underline.Value ? Underline.Single : Underline.None;
        
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
                }
                else
                {
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

        if (indentLevel.HasValue)
        {
            para.ParagraphFormat.LeftIndent = indentLevel.Value * 36;
        }
        else if (leftIndent.HasValue)
        {
            para.ParagraphFormat.LeftIndent = leftIndent.Value;
        }

        if (firstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
        }

        if (tabStops != null && tabStops.Count > 0)
        {
            para.ParagraphFormat.TabStops.Clear();
            
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
        
        if (paragraphIndex.HasValue && targetPara != null)
        {
            if (paragraphIndex.Value == -1)
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
        
        result += warningMessage;

        return await Task.FromResult(result);
    }
}

