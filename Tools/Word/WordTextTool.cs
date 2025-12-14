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
    public string Description => @"Perform text operations in Word documents. Supports 8 operations: add, delete, replace, search, format, insert_at_position, delete_range, add_with_style.

Usage examples:
- Add text: word_text(operation='add', path='doc.docx', text='Hello World')
- Add formatted text: word_text(operation='add', path='doc.docx', text='Bold text', bold=true)
- Replace text: word_text(operation='replace', path='doc.docx', searchText='old', replaceText='new')
- Search text: word_text(operation='search', path='doc.docx', searchText='keyword')
- Format text: word_text(operation='format', path='doc.docx', paragraphIndex=0, runIndex=0, bold=true)
- Insert at position: word_text(operation='insert_at_position', path='doc.docx', paragraphIndex=0, runIndex=0, text='Inserted')
- Delete text: word_text(operation='delete', path='doc.docx', searchText='text to delete')
- Delete range: word_text(operation='delete_range', path='doc.docx', startParagraphIndex=0, startRunIndex=0, endParagraphIndex=0, endRunIndex=5)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add text at document end (required params: path, text)
- 'delete': Delete text matching searchText (required params: path, searchText)
- 'replace': Replace text (required params: path, searchText, replaceText)
- 'search': Search for text (required params: path, searchText)
- 'format': Format existing text (required params: path, paragraphIndex, runIndex)
- 'insert_at_position': Insert text at specific position (required params: path, paragraphIndex, runIndex, text)
- 'delete_range': Delete text range (required params: path, startParagraphIndex, startRunIndex, endParagraphIndex, endRunIndex)
- 'add_with_style': Add text with style (required params: path, text, styleName)",
                @enum = new[] { "add", "delete", "replace", "search", "format", "insert_at_position", "delete_range", "add_with_style" }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
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
                description = "Underline style: none, single, double, dotted, dash (optional, for add/format operations)",
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
                description = "Strikethrough (optional, for add/format operations)"
            },
            superscript = new
            {
                type = "boolean",
                description = "Superscript (optional, for add/format operations)"
            },
            subscript = new
            {
                type = "boolean",
                description = "Subscript (optional, for add/format operations)"
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
            replaceInFields = new
            {
                type = "boolean",
                description = "Replace text inside fields (optional, default: false). If false, fields like hyperlinks will be excluded from replacement to preserve their functionality"
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
                description = "Section index (0-based, optional, default: 0, for format/insert_at_position/delete_range operations)"
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
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
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

    /// <summary>
    /// Adds text to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing text, optional fontName, fontSize, fontColor, formatting options</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        var text = ArgumentHelper.GetString(arguments, "text", "text");
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double>();
        var bold = arguments?["bold"]?.GetValue<bool>() ?? false;
        var italic = arguments?["italic"]?.GetValue<bool>() ?? false;
        var underline = arguments?["underline"]?.GetValue<string>();
        var color = arguments?["color"]?.GetValue<string>();
        var strikethrough = arguments?["strikethrough"]?.GetValue<bool>() ?? false;
        var superscript = arguments?["superscript"]?.GetValue<bool>() ?? false;
        var subscript = arguments?["subscript"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        
        // Get the last section's body to append the new paragraph
        doc.EnsureMinimum();
        var lastSection = doc.LastSection;
        var body = lastSection.Body;
        
        // Only split text if it actually contains newlines
        // If text doesn't contain newlines, treat it as a single paragraph
        // This prevents creating unnecessary paragraphs that cause format misapplication
        var lines = text.Contains('\n') || text.Contains('\r') 
            ? text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)
            : new[] { text };
        
        var builder = new DocumentBuilder(doc);
        
        // Move to last paragraph in document body (not inside Shape/TextBox)
        // MoveToDocumentEnd() might move inside textbox if cursor is already there
        var bodyParagraphs = body.GetChildNodes(NodeType.Paragraph, false);
        if (bodyParagraphs.Count > 0)
        {
            var lastBodyPara = bodyParagraphs[bodyParagraphs.Count - 1] as Paragraph;
            if (lastBodyPara != null)
            {
                builder.MoveTo(lastBodyPara);
            }
            else
            {
                builder.MoveToDocumentEnd();
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
        }
        
        // Ensure we're in document body, not inside Shape/TextBox
        var currentNode = builder.CurrentNode;
        if (currentNode != null)
        {
            var shapeAncestor = currentNode.GetAncestor(NodeType.Shape);
            if (shapeAncestor != null)
            {
                bodyParagraphs = body.GetChildNodes(NodeType.Paragraph, false);
                if (bodyParagraphs.Count > 0)
                {
                    var lastBodyPara = bodyParagraphs[bodyParagraphs.Count - 1] as Paragraph;
                    if (lastBodyPara != null)
                    {
                        builder.MoveTo(lastBodyPara);
                    }
                }
                else
                {
                    // No paragraphs, append to body directly
                    builder.MoveTo(body);
                }
            }
        }
        
        var createdRuns = new List<Run>();
        
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i];
            
            var currentParaBefore = builder.CurrentParagraph;
            bool needsNewParagraph = false;
            if (currentParaBefore != null)
            {
                var existingRuns = currentParaBefore.GetChildNodes(NodeType.Run, false);
                var existingText = currentParaBefore.GetText().Trim();
                needsNewParagraph = existingRuns.Count > 0 || !string.IsNullOrEmpty(existingText);
            }
            
            if (needsNewParagraph || i == 0)
            {
                builder.Writeln();
                builder.MoveTo(builder.CurrentParagraph);
            }
            
            builder.Font.ClearFormatting();
            builder.Font.Bold = false;
            builder.Font.Italic = false;
            builder.Font.Underline = Underline.None;
            builder.Font.StrikeThrough = false;
            builder.Font.Superscript = false;
            builder.Font.Subscript = false;
            builder.ParagraphFormat.ClearFormatting();
            
            if (!string.IsNullOrEmpty(fontName))
                builder.Font.Name = fontName;
            
            if (fontSize.HasValue)
                builder.Font.Size = fontSize.Value;
            
            if (arguments?["bold"] != null)
                builder.Font.Bold = bold;
            
            if (arguments?["italic"] != null)
                builder.Font.Italic = italic;
            
            if (!string.IsNullOrEmpty(underline))
            {
                var underlineValue = underline.ToLower() switch
                {
                    "single" => Underline.Single,
                    "double" => Underline.Double,
                    "dotted" => Underline.Dotted,
                    "dash" => Underline.Dash,
                    "none" => Underline.None,
                    _ => Underline.None
                };
                builder.Font.Underline = underlineValue;
            }
            
            if (arguments?["strikethrough"] != null)
                builder.Font.StrikeThrough = strikethrough;
            
            if (arguments?["superscript"] != null)
            {
                builder.Font.Subscript = false;
                builder.Font.Superscript = superscript;
            }
            else if (arguments?["subscript"] != null)
            {
                builder.Font.Superscript = false;
                builder.Font.Subscript = subscript;
            }
            
            if (!string.IsNullOrEmpty(color))
            {
                try
                {
                    builder.Font.Color = ColorHelper.ParseColor(color);
                }
                catch
                {
                    // Ignore invalid color
                }
            }
            
            // Get the current paragraph and its runs before writing
            var currentPara = builder.CurrentParagraph;
            int runsBefore = 0;
            if (currentPara != null)
            {
                runsBefore = currentPara.GetChildNodes(NodeType.Run, false).Count;
            }
            
            // Write text using DocumentBuilder - this ensures format is applied correctly
            builder.Write(line);
            
            // Get ALL runs in the paragraph and ensure format is applied to ALL of them
            // This includes any runs that might have been created by DocumentBuilder internally
            if (currentPara != null)
            {
                var runs = currentPara.GetChildNodes(NodeType.Run, false);
                int runsAfter = runs.Count;
                
                for (int r = runsBefore; r < runsAfter; r++)
                {
                    var run = runs[r] as Run;
                    if (run != null)
                    {
                        bool isNewRun = r >= runsBefore;
                        bool textMatches = run.Text == line;
                        
                        if (isNewRun && textMatches)
                        {
                            run.Font.Subscript = false;
                            run.Font.Superscript = false;
                            run.Font.StrikeThrough = false;
                            run.Font.Bold = false;
                            run.Font.Italic = false;
                            run.Font.Underline = Underline.None;
                            
                            // Re-apply format to ensure it's applied correctly
                            if (arguments?["bold"] != null)
                                run.Font.Bold = bold;
                            
                            if (arguments?["italic"] != null)
                                run.Font.Italic = italic;
                            
                            if (!string.IsNullOrEmpty(underline))
                            {
                                var underlineValue = underline.ToLower() switch
                                {
                                    "single" => Underline.Single,
                                    "double" => Underline.Double,
                                    "dotted" => Underline.Dotted,
                                    "dash" => Underline.Dash,
                                    "none" => Underline.None,
                                    _ => Underline.None
                                };
                                run.Font.Underline = underlineValue;
                            }
                            
                            if (arguments?["strikethrough"] != null)
                                run.Font.StrikeThrough = strikethrough;
                            
                            if (arguments?["superscript"] != null)
                            {
                                run.Font.Subscript = false;
                                run.Font.Superscript = superscript;
                            }
                            else if (arguments?["subscript"] != null)
                            {
                                run.Font.Superscript = false;
                                run.Font.Subscript = subscript;
                            }
                            
                            createdRuns.Add(run);
                        }
                    }
                }
            }
        }
        
        // Save document
        doc.Save(outputPath);

        var formatInfo = new List<string>();
        if (bold) formatInfo.Add("粗體");
        if (italic) formatInfo.Add("斜體");
        if (!string.IsNullOrEmpty(underline) && underline != "none") formatInfo.Add($"底線({underline})");
        if (strikethrough) formatInfo.Add("刪除線");
        if (superscript) formatInfo.Add("上標");
        if (subscript) formatInfo.Add("下標");
        
        var result = $"成功添加文字到文檔\n";
        if (formatInfo.Count > 0)
        {
            result += $"應用格式: {string.Join(", ", formatInfo)}\n";
        }
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    /// Deletes text from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing searchText, optional matchCase, matchWholeWord, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message with deletion count</returns>
    private async Task<string> DeleteTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        var searchText = arguments?["searchText"]?.GetValue<string>();
        var startParagraphIndex = arguments?["startParagraphIndex"]?.GetValue<int?>();
        var startRunIndex = arguments?["startRunIndex"]?.GetValue<int>() ?? 0;
        var endParagraphIndex = arguments?["endParagraphIndex"]?.GetValue<int?>();
        var endRunIndex = arguments?["endRunIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        
        // If searchText is provided, find the text and determine paragraph/run indices
        if (!string.IsNullOrEmpty(searchText))
        {
            // Search for the text in the document
            bool found = false;
            for (int p = 0; p < paragraphs.Count; p++)
            {
                var para = paragraphs[p] as Paragraph;
                if (para == null) continue;
                
                var paraText = para.GetText();
                int textIndex = paraText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);
                
                if (textIndex >= 0)
                {
                    // Found the text, determine run indices
                    var runs = para.GetChildNodes(NodeType.Run, false);
                    int charCount = 0;
                    int startRunIdx = 0;
                    int endRunIdx = runs.Count - 1;
                    
                    // Find the run containing the start of searchText
                    for (int r = 0; r < runs.Count; r++)
                    {
                        var run = runs[r] as Run;
                        if (run == null) continue;
                        
                        int runLength = run.Text.Length;
                        if (charCount + runLength > textIndex)
                        {
                            startRunIdx = r;
                            break;
                        }
                        charCount += runLength;
                    }
                    
                    // Find the run containing the end of searchText
                    charCount = 0;
                    int endTextIndex = textIndex + searchText.Length;
                    for (int r = 0; r < runs.Count; r++)
                    {
                        var run = runs[r] as Run;
                        if (run == null) continue;
                        
                        int runLength = run.Text.Length;
                        if (charCount + runLength >= endTextIndex)
                        {
                            endRunIdx = r;
                            break;
                        }
                        charCount += runLength;
                    }
                    
                    startParagraphIndex = p;
                    endParagraphIndex = p;
                    startRunIndex = startRunIdx;
                    endRunIndex = endRunIdx;
                    found = true;
                    break;
                }
            }
            
            if (!found)
            {
                throw new ArgumentException($"未找到文字 '{searchText}'。請使用 search 操作先確認文字位置。");
            }
        }
        else
        {
            // Require paragraph indices if searchText is not provided
            if (!startParagraphIndex.HasValue)
                throw new ArgumentException("startParagraphIndex is required when searchText is not provided");
            if (!endParagraphIndex.HasValue)
                throw new ArgumentException("endParagraphIndex is required when searchText is not provided");
        }
        
        if (!startParagraphIndex.HasValue || !endParagraphIndex.HasValue)
        {
            throw new ArgumentException("無法確定段落索引");
        }
        
        if (startParagraphIndex.Value < 0 || startParagraphIndex.Value >= paragraphs.Count ||
            endParagraphIndex.Value < 0 || endParagraphIndex.Value >= paragraphs.Count ||
            startParagraphIndex.Value > endParagraphIndex.Value)
        {
            throw new ArgumentException($"段落索引超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
        }
        
        var startPara = paragraphs[startParagraphIndex.Value] as Paragraph;
        var endPara = paragraphs[endParagraphIndex.Value] as Paragraph;
        
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
            
            if (startParagraphIndex.Value == endParagraphIndex.Value)
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
                
                for (int p = startParagraphIndex.Value + 1; p < endParagraphIndex.Value; p++)
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
        if (startParagraphIndex.Value == endParagraphIndex.Value)
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
            
            for (int p = endParagraphIndex.Value - 1; p > startParagraphIndex.Value; p--)
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
        if (!string.IsNullOrEmpty(searchText))
        {
            result += $"刪除文字: {searchText}\n";
        }
        result += $"範圍: 段落 {startParagraphIndex.Value} Run {startRunIndex} 到 段落 {endParagraphIndex.Value} Run {endRunIndex ?? -1}\n";
        if (!string.IsNullOrEmpty(preview))
        {
            result += $"刪除內容預覽: {preview}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Replaces text in the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing searchText, replaceText, optional matchCase, matchWholeWord, outputPath</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message with replacement count</returns>
    private async Task<string> ReplaceTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        var find = ArgumentHelper.GetString(arguments, "find", "find");
        var replace = ArgumentHelper.GetString(arguments, "replace", "replace");
        var useRegex = arguments?["useRegex"]?.GetValue<bool>() ?? false;
        var replaceInFields = arguments?["replaceInFields"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        
        var options = new FindReplaceOptions();
        
        // Fields (like hyperlinks) should not be replaced unless explicitly requested
        if (!replaceInFields)
        {
            options.ReplacingCallback = new FieldSkipReplacingCallback();
        }
        
        if (useRegex)
        {
            doc.Range.Replace(new Regex(find), replace, options);
        }
        else
        {
            doc.Range.Replace(find, replace, options);
        }

        doc.Save(outputPath);

        var result = $"Text replaced in document: {outputPath}";
        if (!replaceInFields)
        {
            result += "\nNote: Fields (such as hyperlinks) were excluded from replacement to preserve their functionality.";
        }
        return await Task.FromResult(result);
    }

    /// <summary>
    /// Searches for text in the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing searchText, optional matchCase, matchWholeWord</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with search results</returns>
    private async Task<string> SearchTextAsync(JsonObject? arguments, string path)
    {
        var searchText = ArgumentHelper.GetString(arguments, "searchText", "searchText");
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

    /// <summary>
    /// Formats text in the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing searchText, optional formatting options, matchCase, matchWholeWord</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message with format count</returns>
    private async Task<string> FormatTextAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex", "paragraphIndex");
        var runIndex = arguments?["runIndex"]?.GetValue<int?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>() ?? 0;
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
        
        // Use section-based paragraph indexing to match GetRunFormat behavior
        // This ensures consistent indexing between format and get operations
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex {sectionIndex} 超出範圍 (文檔共有 {doc.Sections.Count} 個節)");
        }
        
        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"段落索引 {paragraphIndex} 超出範圍 (節 {sectionIndex} 的正文共有 {paragraphs.Count} 個段落)");
        }
        
        var para = paragraphs[paragraphIndex];
        
        var runs = para.GetChildNodes(NodeType.Run, false);
        if (runs == null || runs.Count == 0)
        {
            var newRun = new Run(doc, "");
            para.AppendChild(newRun);
            runs = para.GetChildNodes(NodeType.Run, false);
            if (runs == null || runs.Count == 0)
            {
                throw new InvalidOperationException($"段落 #{paragraphIndex} 中沒有 Run 節點，且無法創建新的 Run 節點");
            }
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
            // Clear conflicting formats before applying new ones to prevent accumulation
            // This ensures formats are applied correctly without interference from previous formats
            
            // Clear superscript/subscript if either is being set (they are mutually exclusive)
            if (superscript.HasValue || subscript.HasValue)
            {
                if (superscript.HasValue && superscript.Value)
                {
                    run.Font.Subscript = false; // Clear subscript when setting superscript
                }
                else if (subscript.HasValue && subscript.Value)
                {
                    run.Font.Superscript = false; // Clear superscript when setting subscript
                }
                else
                {
                    // If setting to false, clear both
                    if (superscript.HasValue && !superscript.Value)
                    {
                        run.Font.Superscript = false;
                    }
                    if (subscript.HasValue && !subscript.Value)
                    {
                        run.Font.Subscript = false;
                    }
                }
            }
            
            // Clear underline if it's being set (to avoid conflicts)
            if (!string.IsNullOrEmpty(underline))
            {
                // Will be set below, but ensure no other underline styles interfere
            }
            
            // Apply font names
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
            
            // Apply bold/italic - ensure they are set correctly
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
            
            // Apply underline - ensure it's set correctly
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
                    run.Font.Color = ColorHelper.ParseColor(color);
                    var colorValue = color.TrimStart('#');
                    changes.Add($"顏色: {(colorValue.Length == 6 ? "#" : "")}{colorValue}");
                }
                catch
                {
                    // Ignore color parsing errors
                }
            }
            
            // Apply strikethrough
            if (strikethrough.HasValue)
            {
                run.Font.StrikeThrough = strikethrough.Value;
                changes.Add($"刪除線: {(strikethrough.Value ? "是" : "否")}");
            }
            
            // Apply superscript/subscript (already cleared conflicting ones above)
            if (superscript.HasValue)
            {
                run.Font.Superscript = superscript.Value;
                changes.Add($"上標: {(superscript.Value ? "是" : "否")}");
            }
            
            if (subscript.HasValue)
            {
                run.Font.Subscript = subscript.Value;
                changes.Add($"下標: {(subscript.Value ? "是" : "否")}");
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

    /// <summary>
    /// Inserts text at a specific position
    /// </summary>
    /// <param name="arguments">JSON arguments containing text, paragraphIndex, runIndex, optional formatting options</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> InsertTextAtPositionAsync(JsonObject? arguments, string path, string outputPath)
    {
        var paragraphIndex = ArgumentHelper.GetInt(arguments, "insertParagraphIndex", "insertParagraphIndex");
        var charIndex = ArgumentHelper.GetInt(arguments, "charIndex", "charIndex");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var text = ArgumentHelper.GetString(arguments, "text", "text");
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

    /// <summary>
    /// Deletes text in a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing startParagraphIndex, startRunIndex, endParagraphIndex, endRunIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteTextRangeAsync(JsonObject? arguments, string path, string outputPath)
    {
        var startParagraphIndex = ArgumentHelper.GetInt(arguments, "startParagraphIndex", "startParagraphIndex");
        var startCharIndex = ArgumentHelper.GetInt(arguments, "startCharIndex", "startCharIndex");
        var endParagraphIndex = ArgumentHelper.GetInt(arguments, "endParagraphIndex", "endParagraphIndex");
        var endCharIndex = ArgumentHelper.GetInt(arguments, "endCharIndex", "endCharIndex");
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

    /// <summary>
    /// Adds text with a specific style
    /// </summary>
    /// <param name="arguments">JSON arguments containing text, styleName, optional formatting options</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private async Task<string> AddTextWithStyleAsync(JsonObject? arguments, string path, string outputPath)
    {
        var text = ArgumentHelper.GetString(arguments, "text", "text");
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
                run.Font.Color = ColorHelper.ParseColor(color);
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

// Helper class to skip field replacement
internal class FieldSkipReplacingCallback : IReplacingCallback
{
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Skip replacement if we're inside a field
        if (args.MatchNode.GetAncestor(NodeType.FieldStart) != null ||
            args.MatchNode.GetAncestor(NodeType.FieldSeparator) != null ||
            args.MatchNode.GetAncestor(NodeType.FieldEnd) != null)
        {
            return ReplaceAction.Skip;
        }
        return ReplaceAction.Replace;
    }
}

