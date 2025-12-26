using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for text operations in Word documents
///     Merges: WordAddTextTool, WordDeleteTextTool, WordReplaceTextTool, WordSearchTextTool,
///     WordFormatTextTool, WordInsertTextAtPositionTool, WordDeleteTextRangeTool, WordAddTextWithStyleTool
/// </summary>
public class WordTextTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Perform text operations in Word documents. Supports 8 operations: add, delete, replace, search, format, insert_at_position, delete_range, add_with_style.

Usage examples:
- Add text: word_text(operation='add', path='doc.docx', text='Hello World')
- Add formatted text: word_text(operation='add', path='doc.docx', text='Bold text', bold=true)
- Replace text: word_text(operation='replace', path='doc.docx', find='old', replace='new')
- Search text: word_text(operation='search', path='doc.docx', searchText='keyword')
- Format text: word_text(operation='format', path='doc.docx', paragraphIndex=0, runIndex=0, bold=true)
- Insert at position: word_text(operation='insert_at_position', path='doc.docx', paragraphIndex=0, runIndex=0, text='Inserted')
- Delete text: word_text(operation='delete', path='doc.docx', searchText='text to delete') or word_text(operation='delete', path='doc.docx', startParagraphIndex=0, endParagraphIndex=0)
- Delete range: word_text(operation='delete_range', path='doc.docx', startParagraphIndex=0, startRunIndex=0, endParagraphIndex=0, endRunIndex=5)";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
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
- 'delete': Delete text (required params: path, searchText OR startParagraphIndex+endParagraphIndex)
- 'replace': Replace text (required params: path, find, replace)
- 'search': Search for text (required params: path, searchText)
- 'format': Format existing text (required params: path, paragraphIndex, runIndex)
- 'insert_at_position': Insert text at specific position (required params: path, paragraphIndex, runIndex, text)
- 'delete_range': Delete text range (required params: path, startParagraphIndex, startRunIndex, endParagraphIndex, endRunIndex)
- 'add_with_style': Add text with style (required params: path, text, styleName)",
                @enum = new[]
                {
                    "add", "delete", "replace", "search", "format", "insert_at_position", "delete_range",
                    "add_with_style"
                }
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
                description = "Font name (optional, e.g., 'Arial')"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, e.g., 'Times New Roman')"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean, optional)"
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
                description =
                    "Underline style: none, single, double, dotted, dash (optional, for add/format operations)",
                @enum = new[] { "none", "single", "double", "dotted", "dash" }
            },
            color = new
            {
                type = "string",
                description =
                    "Text color (hex format like 'FF0000' or '#FF0000' for red, or name like 'Red', 'Blue', optional)"
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
                description =
                    "Replace text inside fields (optional, default: false). If false, fields like hyperlinks will be excluded from replacement to preserve their functionality"
            },
            // Search parameters
            searchText = new
            {
                type = "string",
                description =
                    "Text to search for (required for search operation, optional for delete operation as alternative to startParagraphIndex+endParagraphIndex)"
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
                description =
                    "Number of characters to show before and after match for context (optional, default: 50, for search operation)"
            },
            // Delete parameters
            startParagraphIndex = new
            {
                type = "number",
                description =
                    "Start paragraph index (0-based, required for delete operation if searchText is not provided)"
            },
            startRunIndex = new
            {
                type = "number",
                description =
                    "Start run index within start paragraph (0-based, optional, default: 0, for delete operation)"
            },
            endParagraphIndex = new
            {
                type = "number",
                description =
                    "End paragraph index (0-based, inclusive, required for delete operation if searchText is not provided)"
            },
            endRunIndex = new
            {
                type = "number",
                description =
                    "End run index within end paragraph (0-based, inclusive, optional, default: last run, for delete operation)"
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
                description =
                    "Run index within the paragraph (0-based, optional, formats all runs if not provided, for format operation)"
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
                description =
                    "Section index (0-based, optional, default: 0, for format/insert_at_position/delete_range operations)"
            },
            insertBefore = new
            {
                type = "boolean",
                description =
                    "Insert before position (optional, default: false, inserts after, for insert_at_position operation)"
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
                description =
                    "Style name to apply (e.g., 'Heading 1', 'Normal', optional, for add_with_style operation)"
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
                description =
                    "Indentation level (0-8, where each level = 36 points / 0.5 inch, optional, for add_with_style operation)"
            },
            leftIndent = new
            {
                type = "number",
                description = "Left indentation in points (optional, for add_with_style operation)"
            },
            firstLineIndent = new
            {
                type = "number",
                description =
                    "First line indentation in points (positive = indent first line, negative = hanging indent, optional, for add_with_style operation)"
            },
            paragraphIndexForAdd = new
            {
                type = "number",
                description =
                    "Index of the paragraph to insert after (0-based, optional, for add_with_style operation). Use -1 to insert at the beginning."
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
                        alignment = new
                        {
                            type = "string", description = "Tab alignment: Left, Center, Right, Decimal, Bar",
                            @enum = new[] { "Left", "Center", "Right", "Decimal", "Bar" }
                        },
                        leader = new
                        {
                            type = "string", description = "Tab leader: None, Dots, Dashes, Line, Heavy, MiddleDot",
                            @enum = new[] { "None", "Dots", "Dashes", "Line", "Heavy", "MiddleDot" }
                        }
                    }
                }
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation switch
        {
            "add" => await AddTextAsync(path, outputPath, arguments),
            "delete" => await DeleteTextAsync(path, outputPath, arguments),
            "replace" => await ReplaceTextAsync(path, outputPath, arguments),
            "search" => await SearchTextAsync(path, arguments),
            "format" => await FormatTextAsync(path, outputPath, arguments),
            "insert_at_position" => await InsertTextAtPositionAsync(path, outputPath, arguments),
            "delete_range" => await DeleteTextRangeAsync(path, outputPath, arguments),
            "add_with_style" => await AddTextWithStyleAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds text to the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, optional fontName, fontSize, fontColor, formatting options</param>
    /// <returns>Success message</returns>
    private Task<string> AddTextAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var text = ArgumentHelper.GetString(arguments, "text");
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBool(arguments, "bold", false);
            var italic = ArgumentHelper.GetBool(arguments, "italic", false);
            var underline = ArgumentHelper.GetStringNullable(arguments, "underline");
            var color = ArgumentHelper.GetStringNullable(arguments, "color");
            var strikethrough = ArgumentHelper.GetBool(arguments, "strikethrough", false);
            var superscript = ArgumentHelper.GetBool(arguments, "superscript", false);
            var subscript = ArgumentHelper.GetBool(arguments, "subscript", false);

            var doc = new Document(path);

            doc.EnsureMinimum();
            var lastSection = doc.LastSection;
            var body = lastSection.Body;

            // Split text by newlines to create multiple paragraphs, preventing format misapplication
            var lines = text.Contains('\n') || text.Contains('\r')
                ? text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None)
                : [text];

            var builder = new DocumentBuilder(doc);

            // Move to last paragraph in document body (not inside Shape/TextBox)
            var bodyParagraphs = body.GetChildNodes(NodeType.Paragraph, false);
            if (bodyParagraphs.Count > 0)
            {
                if (bodyParagraphs[^1] is Paragraph lastBodyPara)
                    builder.MoveTo(lastBodyPara);
                else
                    builder.MoveToDocumentEnd();
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
                        if (bodyParagraphs[^1] is Paragraph lastBodyPara) builder.MoveTo(lastBodyPara);
                    }
                    else
                    {
                        builder.MoveTo(body);
                    }
                }
            }

            for (var i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var currentParaBefore = builder.CurrentParagraph;
                var needsNewParagraph = false;
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

                bool? strikethroughValue = arguments?["strikethrough"] != null ? strikethrough : null;
                bool? superscriptValue = arguments?["superscript"] != null ? superscript : null;
                bool? subscriptValue = arguments?["subscript"] != null ? subscript : null;

                FontHelper.Word.ApplyFontSettings(
                    builder,
                    fontName,
                    fontSize: fontSize,
                    bold: arguments?["bold"] != null ? bold : null,
                    italic: arguments?["italic"] != null ? italic : null,
                    underline: underline,
                    color: color,
                    strikethrough: strikethroughValue,
                    superscript: superscriptValue,
                    subscript: subscriptValue
                );

                var currentPara = builder.CurrentParagraph;
                var runsBefore = 0;
                if (currentPara != null) runsBefore = currentPara.GetChildNodes(NodeType.Run, false).Count;

                // Write text using DocumentBuilder to ensure format is applied correctly
                builder.Write(line);

                // Apply format to all runs created by DocumentBuilder
                if (currentPara != null)
                {
                    var runs = currentPara.GetChildNodes(NodeType.Run, false);
                    var runsAfter = runs.Count;

                    for (var r = runsBefore; r < runsAfter; r++)
                        if (runs[r] is Run run)
                        {
                            var isNewRun = r >= runsBefore;
                            var textMatches = run.Text == line;

                            if (isNewRun && textMatches)
                            {
                                run.Font.Subscript = false;
                                run.Font.Superscript = false;
                                run.Font.StrikeThrough = false;
                                run.Font.Bold = false;
                                run.Font.Italic = false;
                                run.Font.Underline = Underline.None;

                                FontHelper.Word.ApplyFontSettings(
                                    run,
                                    fontSize: null,
                                    bold: arguments?["bold"] != null ? bold : null,
                                    italic: arguments?["italic"] != null ? italic : null,
                                    underline: underline,
                                    strikethrough: arguments?["strikethrough"] != null ? strikethrough : null,
                                    superscript: arguments?["superscript"] != null ? superscript : null,
                                    subscript: arguments?["subscript"] != null ? subscript : null
                                );
                            }
                        }
                }
            }

            doc.Save(outputPath);

            var formatInfo = new List<string>();
            if (bold) formatInfo.Add("bold");
            if (italic) formatInfo.Add("italic");
            if (!string.IsNullOrEmpty(underline) && underline != "none") formatInfo.Add($"underline({underline})");
            if (strikethrough) formatInfo.Add("strikethrough");
            if (superscript) formatInfo.Add("superscript");
            if (subscript) formatInfo.Add("subscript");

            var result = "Text added to document successfully\n";
            if (formatInfo.Count > 0) result += $"Applied formats: {string.Join(", ", formatInfo)}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes text from the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing searchText, optional matchCase, matchWholeWord, outputPath</param>
    /// <returns>Success message with deletion count</returns>
    private Task<string> DeleteTextAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var searchText = ArgumentHelper.GetStringNullable(arguments, "searchText");
            var startParagraphIndex = ArgumentHelper.GetIntNullable(arguments, "startParagraphIndex");
            var startRunIndex = ArgumentHelper.GetInt(arguments, "startRunIndex", 0);
            var endParagraphIndex = ArgumentHelper.GetIntNullable(arguments, "endParagraphIndex");
            var endRunIndex = ArgumentHelper.GetIntNullable(arguments, "endRunIndex");

            var doc = new Document(path);
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (!string.IsNullOrEmpty(searchText))
            {
                var found = false;
                for (var p = 0; p < paragraphs.Count; p++)
                {
                    if (paragraphs[p] is not Paragraph para) continue;

                    var paraText = para.GetText();
                    var textIndex = paraText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);

                    if (textIndex >= 0)
                    {
                        var runs = para.GetChildNodes(NodeType.Run, false);
                        var charCount = 0;
                        var startRunIdx = 0;
                        var endRunIdx = runs.Count - 1;

                        for (var r = 0; r < runs.Count; r++)
                        {
                            if (runs[r] is not Run run) continue;

                            var runLength = run.Text.Length;
                            if (charCount + runLength > textIndex)
                            {
                                startRunIdx = r;
                                break;
                            }

                            charCount += runLength;
                        }

                        charCount = 0;
                        var endTextIndex = textIndex + searchText.Length;
                        for (var r = 0; r < runs.Count; r++)
                        {
                            if (runs[r] is not Run run) continue;

                            var runLength = run.Text.Length;
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
                    throw new ArgumentException(
                        $"Text '{searchText}' not found. Please use search operation to confirm text location first.");
            }
            else
            {
                if (!startParagraphIndex.HasValue)
                    throw new ArgumentException("startParagraphIndex is required when searchText is not provided");
                if (!endParagraphIndex.HasValue)
                    throw new ArgumentException("endParagraphIndex is required when searchText is not provided");
            }

            if (!startParagraphIndex.HasValue || !endParagraphIndex.HasValue)
                throw new ArgumentException("Unable to determine paragraph index");

            if (startParagraphIndex.Value < 0 || startParagraphIndex.Value >= paragraphs.Count ||
                endParagraphIndex.Value < 0 || endParagraphIndex.Value >= paragraphs.Count ||
                startParagraphIndex.Value > endParagraphIndex.Value)
                throw new ArgumentException(
                    $"Paragraph index is out of range (document has {paragraphs.Count} paragraphs)");

            if (paragraphs[startParagraphIndex.Value] is not Paragraph startPara ||
                paragraphs[endParagraphIndex.Value] is not Paragraph endPara)
                throw new InvalidOperationException("Unable to find specified paragraph");

            var deletedText = "";
            try
            {
                var startRuns = startPara.GetChildNodes(NodeType.Run, false);
                var endRuns = endPara.GetChildNodes(NodeType.Run, false);

                if (startParagraphIndex.Value == endParagraphIndex.Value)
                {
                    if (startRuns is { Count: > 0 })
                    {
                        var actualEndRunIndex = endRunIndex ?? startRuns.Count - 1;
                        if (startRunIndex >= 0 && startRunIndex < startRuns.Count &&
                            actualEndRunIndex >= 0 && actualEndRunIndex < startRuns.Count &&
                            startRunIndex <= actualEndRunIndex)
                            for (var i = startRunIndex; i <= actualEndRunIndex; i++)
                                if (startRuns[i] is Run run)
                                    deletedText += run.Text;
                    }
                }
                else
                {
                    if (startRuns != null && startRuns.Count > startRunIndex)
                        for (var i = startRunIndex; i < startRuns.Count; i++)
                            if (startRuns[i] is Run run)
                                deletedText += run.Text;

                    for (var p = startParagraphIndex.Value + 1; p < endParagraphIndex.Value; p++)
                        if (paragraphs[p] is Paragraph para)
                            deletedText += para.GetText();

                    if (endRuns is { Count: > 0 })
                    {
                        var actualEndRunIndex = endRunIndex ?? endRuns.Count - 1;
                        for (var i = 0; i <= actualEndRunIndex && i < endRuns.Count; i++)
                            if (endRuns[i] is Run run)
                                deletedText += run.Text;
                    }
                }
            }
            catch (Exception ex)
            {
                // Ignore exceptions when extracting deleted text - this is for informational purposes only
                Console.Error.WriteLine($"[WARN] Error extracting deleted text (informational only): {ex.Message}");
            }

            if (startParagraphIndex.Value == endParagraphIndex.Value)
            {
                var runs = startPara.GetChildNodes(NodeType.Run, false);
                if (runs is { Count: > 0 })
                {
                    var actualEndRunIndex = endRunIndex ?? runs.Count - 1;
                    if (startRunIndex >= 0 && startRunIndex < runs.Count &&
                        actualEndRunIndex >= 0 && actualEndRunIndex < runs.Count &&
                        startRunIndex <= actualEndRunIndex)
                        for (var i = actualEndRunIndex; i >= startRunIndex; i--)
                            runs[i]?.Remove();
                }
            }
            else
            {
                var startRuns = startPara.GetChildNodes(NodeType.Run, false);
                if (startRuns != null && startRuns.Count > startRunIndex)
                    for (var i = startRuns.Count - 1; i >= startRunIndex; i--)
                        startRuns[i]?.Remove();

                for (var p = endParagraphIndex.Value - 1; p > startParagraphIndex.Value; p--) paragraphs[p]?.Remove();

                var endRuns = endPara.GetChildNodes(NodeType.Run, false);
                if (endRuns is { Count: > 0 })
                {
                    var actualEndRunIndex = endRunIndex ?? endRuns.Count - 1;
                    for (var i = actualEndRunIndex; i >= 0; i--)
                        if (i < endRuns.Count)
                            endRuns[i]?.Remove();
                }
            }

            doc.Save(outputPath);

            var preview = deletedText.Length > 50 ? deletedText.Substring(0, 50) + "..." : deletedText;

            var result = "Text deleted successfully\n";
            if (!string.IsNullOrEmpty(searchText)) result += $"Deleted text: {searchText}\n";
            result +=
                $"Range: Paragraph {startParagraphIndex.Value} Run {startRunIndex} to Paragraph {endParagraphIndex.Value} Run {endRunIndex ?? -1}\n";
            if (!string.IsNullOrEmpty(preview)) result += $"Deleted content preview: {preview}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Replaces text in the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">
    ///     JSON arguments containing searchText, replaceText, optional matchCase, matchWholeWord,
    ///     outputPath
    /// </param>
    /// <returns>Success message with replacement count</returns>
    private Task<string> ReplaceTextAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var find = ArgumentHelper.GetString(arguments, "find");
            var replace = ArgumentHelper.GetString(arguments, "replace");
            var useRegex = ArgumentHelper.GetBool(arguments, "useRegex", false);
            var replaceInFields = ArgumentHelper.GetBool(arguments, "replaceInFields", false);

            var doc = new Document(path);

            var options = new FindReplaceOptions();

            // Fields (like hyperlinks) should not be replaced unless explicitly requested
            if (!replaceInFields) options.ReplacingCallback = new FieldSkipReplacingCallback();

            if (useRegex)
                doc.Range.Replace(new Regex(find), replace, options);
            else
                doc.Range.Replace(find, replace, options);

            doc.Save(outputPath);

            var result = $"Text replaced in document: {outputPath}";
            if (!replaceInFields)
                result +=
                    "\nNote: Fields (such as hyperlinks) were excluded from replacement to preserve their functionality.";
            return result;
        });
    }

    /// <summary>
    ///     Searches for text in the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="arguments">JSON arguments containing searchText, optional matchCase, matchWholeWord</param>
    /// <returns>Formatted string with search results</returns>
    private Task<string> SearchTextAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var searchText = ArgumentHelper.GetString(arguments, "searchText");
            var useRegex = ArgumentHelper.GetBool(arguments, "useRegex", false);
            var caseSensitive = ArgumentHelper.GetBool(arguments, "caseSensitive", false);
            var maxResults = ArgumentHelper.GetInt(arguments, "maxResults", 50);
            var contextLength = ArgumentHelper.GetInt(arguments, "contextLength", 50);

            var doc = new Document(path);
            var result = new StringBuilder();
            var matches = new List<(string text, int paragraphIndex, string context)>();

            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            for (var i = 0; i < paragraphs.Count && matches.Count < maxResults; i++)
            {
                if (paragraphs[i] is not Paragraph para) continue;

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
                    var index = 0;

                    while ((index = paraText.IndexOf(searchText, index, comparison)) != -1)
                    {
                        if (matches.Count >= maxResults) break;

                        var context = GetContext(paraText, index, searchText.Length, contextLength);
                        matches.Add((searchText, i, context));
                        index += searchText.Length;
                    }
                }
            }

            result.AppendLine("=== Search Results ===");
            result.AppendLine($"Search text: {searchText}");
            result.AppendLine($"Use regex: {(useRegex ? "Yes" : "No")}");
            result.AppendLine($"Case sensitive: {(caseSensitive ? "Yes" : "No")}");
            result.AppendLine(
                $"Found {matches.Count} matches{(matches.Count >= maxResults ? $" (limited to first {maxResults})" : "")}\n");

            if (matches.Count == 0)
                result.AppendLine("No matching text found");
            else
                for (var i = 0; i < matches.Count; i++)
                {
                    var match = matches[i];
                    result.AppendLine($"Match #{i + 1}:");
                    result.AppendLine($"  Location: Paragraph #{match.paragraphIndex}");
                    result.AppendLine($"  Matched text: {match.text}");
                    result.AppendLine($"  Context: ...{match.context}...");
                    result.AppendLine();
                }

            return result.ToString();
        });
    }

    /// <summary>
    ///     Extracts context around a matched text for search results display
    /// </summary>
    /// <param name="text">The full paragraph text</param>
    /// <param name="matchIndex">Starting index of the match</param>
    /// <param name="matchLength">Length of the matched text</param>
    /// <param name="contextLength">Number of characters to include before and after the match</param>
    /// <returns>Context string with the match highlighted using 【】 brackets</returns>
    private string GetContext(string text, int matchIndex, int matchLength, int contextLength)
    {
        var start = Math.Max(0, matchIndex - contextLength);
        var end = Math.Min(text.Length, matchIndex + matchLength + contextLength);

        var context = text.Substring(start, end - start);

        context = context.Replace("\r", "").Replace("\n", " ").Trim();

        var highlightStart = matchIndex - start;
        var highlightEnd = highlightStart + matchLength;

        if (highlightStart >= 0 && highlightEnd <= context.Length)
            context = context.Substring(0, highlightStart) +
                      "【" + context.Substring(highlightStart, matchLength) + "】" +
                      context.Substring(highlightEnd);

        return context;
    }

    /// <summary>
    ///     Formats text in the document
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing searchText, optional formatting options, matchCase, matchWholeWord</param>
    /// <returns>Success message with format count</returns>
    private Task<string> FormatTextAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "paragraphIndex");
            var runIndex = ArgumentHelper.GetIntNullable(arguments, "runIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
            var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
            var underline = ArgumentHelper.GetStringNullable(arguments, "underline");
            var color = ArgumentHelper.GetStringNullable(arguments, "color");
            var strikethrough = ArgumentHelper.GetBoolNullable(arguments, "strikethrough");
            var superscript = ArgumentHelper.GetBoolNullable(arguments, "superscript");
            var subscript = ArgumentHelper.GetBoolNullable(arguments, "subscript");

            var doc = new Document(path);

            // Use section-based paragraph indexing to match GetRunFormat behavior
            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException(
                    $"sectionIndex {sectionIndex} is out of range (document has {doc.Sections.Count} sections)");

            var section = doc.Sections[sectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex} is out of range (section {sectionIndex} body has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paragraphIndex];

            var runs = para.GetChildNodes(NodeType.Run, false);
            if (runs == null || runs.Count == 0)
            {
                var newRun = new Run(doc, "");
                para.AppendChild(newRun);
                runs = para.GetChildNodes(NodeType.Run, false);
                if (runs == null || runs.Count == 0)
                    throw new InvalidOperationException(
                        $"Paragraph #{paragraphIndex} has no Run nodes and cannot create new Run node");
            }

            var changes = new List<string>();
            var runsToFormat = new List<Run>();

            if (runIndex.HasValue)
            {
                if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
                    throw new ArgumentException(
                        $"Run index {runIndex.Value} is out of range (paragraph has {runs.Count} Runs)");
                if (runs[runIndex.Value] is Run run)
                    runsToFormat.Add(run);
            }
            else
            {
                foreach (var node in runs)
                    if (node is Run run)
                        runsToFormat.Add(run);
            }

            foreach (var run in runsToFormat)
            {
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
                        if (superscript.HasValue && !superscript.Value) run.Font.Superscript = false;
                        if (subscript.HasValue && !subscript.Value) run.Font.Subscript = false;
                    }
                }

                FontHelper.Word.ApplyFontSettings(
                    run,
                    fontName,
                    fontNameAscii,
                    fontNameFarEast,
                    fontSize,
                    bold,
                    italic,
                    underline,
                    color,
                    strikethrough,
                    superscript,
                    subscript
                );

                if (!string.IsNullOrEmpty(fontNameAscii))
                    changes.Add($"Font (ASCII): {fontNameAscii}");

                if (!string.IsNullOrEmpty(fontNameFarEast))
                    changes.Add($"Font (Far East): {fontNameFarEast}");

                if (!string.IsNullOrEmpty(fontName))
                    if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                        changes.Add($"Font: {fontName}");

                if (fontSize.HasValue)
                    changes.Add($"Font size: {fontSize.Value} points");

                if (bold.HasValue)
                    changes.Add($"Bold: {(bold.Value ? "Yes" : "No")}");

                if (italic.HasValue)
                    changes.Add($"Italic: {(italic.Value ? "Yes" : "No")}");

                if (!string.IsNullOrEmpty(underline))
                    changes.Add($"Underline: {underline}");

                if (!string.IsNullOrEmpty(color))
                {
                    var colorValue = color.TrimStart('#');
                    changes.Add($"Color: {(colorValue.Length == 6 ? "#" : "")}{colorValue}");
                }

                if (strikethrough.HasValue)
                    changes.Add($"Strikethrough: {(strikethrough.Value ? "Yes" : "No")}");

                if (superscript.HasValue)
                    changes.Add($"Superscript: {(superscript.Value ? "Yes" : "No")}");

                if (subscript.HasValue)
                    changes.Add($"Subscript: {(subscript.Value ? "Yes" : "No")}");
            }

            doc.Save(outputPath);

            var result = "Run-level formatting set successfully\n";
            result += $"Paragraph index: {paragraphIndex}\n";
            if (runIndex.HasValue)
                result += $"Run index: {runIndex.Value}\n";
            else
                result += $"Formatted Runs: {runsToFormat.Count}\n";
            if (changes.Count > 0)
                result += $"Changes: {string.Join(", ", changes.Distinct())}\n";
            else
                result += "No change parameters provided\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Inserts text at a specific position
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, paragraphIndex, runIndex, optional formatting options</param>
    /// <returns>Success message</returns>
    private Task<string> InsertTextAtPositionAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var paragraphIndex = ArgumentHelper.GetInt(arguments, "insertParagraphIndex");
            var charIndex = ArgumentHelper.GetInt(arguments, "charIndex");
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");
            var text = ArgumentHelper.GetString(arguments, "text");
            var insertBefore = ArgumentHelper.GetBool(arguments, "insertBefore", false);

            var doc = new Document(path);
            var sectionIdx = sectionIndex ?? 0;
            if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[sectionIdx];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
                throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");

            var para = paragraphs[paragraphIndex];
            var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var totalChars = 0;
            var targetRunIndex = -1;
            var targetRunCharIndex = 0;

            for (var i = 0; i < runs.Count; i++)
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
                if (insertBefore)
                {
                    builder.Write(text);
                }
                else
                {
                    builder.MoveTo(para);
                    builder.MoveToParagraph(paragraphIndex, para.GetText().Length);
                    builder.Write(text);
                }
            }
            else
            {
                var targetRun = runs[targetRunIndex];
                targetRun.Text = targetRun.Text.Insert(targetRunCharIndex, text);
            }

            doc.Save(outputPath);
            return $"Text inserted at position: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes text in a range
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing startParagraphIndex, startRunIndex, endParagraphIndex, endRunIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteTextRangeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var startParagraphIndex = ArgumentHelper.GetInt(arguments, "startParagraphIndex");
            var startCharIndex = ArgumentHelper.GetInt(arguments, "startCharIndex");
            var endParagraphIndex = ArgumentHelper.GetInt(arguments, "endParagraphIndex");
            var endCharIndex = ArgumentHelper.GetInt(arguments, "endCharIndex");
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            var doc = new Document(path);
            var sectionIdx = sectionIndex ?? 0;
            if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[sectionIdx];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count ||
                endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count)
                throw new ArgumentException("Paragraph indices out of range");

            var startPara = paragraphs[startParagraphIndex];
            var endPara = paragraphs[endParagraphIndex];

            if (startParagraphIndex == endParagraphIndex)
            {
                var runs = startPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
                var totalChars = 0;
                int startRunIndex = -1, endRunIndex = -1;
                int startRunCharIndex = 0, endRunCharIndex = 0;

                for (var i = 0; i < runs.Count; i++)
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

                        for (var i = startRunIndex + 1; i < endRunIndex; i++) runs[i].Remove();

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
                    startRun.Text = startRun.Text.Substring(0, startCharIndex);

                for (var i = startParagraphIndex + 1; i < endParagraphIndex; i++) paragraphs[i].Remove();

                var endParaRuns = endPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
                if (endParaRuns.Count > 0 && endCharIndex < endParaRuns[0].Text.Length)
                {
                    endParaRuns[0].Text = endParaRuns[0].Text.Substring(endCharIndex);
                    for (var i = 1; i < endParaRuns.Count; i++) endParaRuns[i].Remove();
                }
            }

            doc.Save(outputPath);
            return $"Text range deleted: {outputPath}";
        });
    }

    /// <summary>
    ///     Adds text with a specific style
    /// </summary>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing text, styleName, optional formatting options</param>
    /// <returns>Success message</returns>
    private Task<string> AddTextWithStyleAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var text = ArgumentHelper.GetString(arguments, "text");
            var styleName = ArgumentHelper.GetStringNullable(arguments, "styleName");
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
            var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
            var underline = ArgumentHelper.GetBoolNullable(arguments, "underline");
            var color = ArgumentHelper.GetStringNullable(arguments, "color");
            var alignment = ArgumentHelper.GetStringNullable(arguments, "alignment");
            var indentLevel = ArgumentHelper.GetIntNullable(arguments, "indentLevel");
            var leftIndent = ArgumentHelper.GetDoubleNullable(arguments, "leftIndent");
            var firstLineIndent = ArgumentHelper.GetDoubleNullable(arguments, "firstLineIndent");
            var tabStops = ArgumentHelper.GetArray(arguments, "tabStops", false);
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndexForAdd");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            Paragraph? targetPara = null;

            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                if (paragraphIndex.Value == -1)
                {
                    if (paragraphs.Count > 0)
                        if (paragraphs[0] is Paragraph firstPara)
                        {
                            targetPara = firstPara;
                            builder.MoveTo(targetPara);
                        }
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    if (paragraphs[paragraphIndex.Value] is Paragraph targetParagraph)
                    {
                        targetPara = targetParagraph;
                        builder.MoveTo(targetPara);
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            $"Unable to find paragraph at index {paragraphIndex.Value}");
                    }
                }
                else
                {
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
            }

            var para = new Paragraph(doc);
            var run = new Run(doc, text);

            var hasCustomParams = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) ||
                                  !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue ||
                                  bold.HasValue || italic.HasValue || underline.HasValue ||
                                  !string.IsNullOrEmpty(color) || !string.IsNullOrEmpty(alignment);

            var warningMessage = "";
            if (!string.IsNullOrEmpty(styleName) && hasCustomParams)
                warningMessage =
                    "\n⚠️ Note: When using both styleName and custom parameters, custom parameters will override corresponding properties in the style.\n" +
                    "This allows you to customize specific properties while applying a style.\n" +
                    "If you need a fully custom style, it is recommended to use word_create_style to create a custom style.\n" +
                    "Example: word_create_style(styleName='Custom Heading', baseStyle='Heading 1', color='000000')";

            if (!string.IsNullOrEmpty(styleName))
                try
                {
                    var style = doc.Styles[styleName];
                    if (style != null)
                        para.ParagraphFormat.StyleName = styleName;
                    else
                        throw new ArgumentException(
                            $"Style '{styleName}' not found. Use word_get_styles tool to view available styles");
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"Unable to apply style '{styleName}': {ex.Message}. Use word_get_styles tool to view available styles",
                        ex);
                }

            var underlineStr = underline.HasValue ? underline.Value ? "single" : "none" : null;
            FontHelper.Word.ApplyFontSettings(
                run,
                fontName,
                fontNameAscii,
                fontNameFarEast,
                fontSize,
                bold,
                italic,
                underlineStr,
                color
            );

            if (!string.IsNullOrEmpty(alignment))
                para.ParagraphFormat.Alignment = alignment.ToLower() switch
                {
                    "left" => ParagraphAlignment.Left,
                    "right" => ParagraphAlignment.Right,
                    "center" => ParagraphAlignment.Center,
                    "justify" => ParagraphAlignment.Justify,
                    _ => ParagraphAlignment.Left
                };

            if (indentLevel.HasValue)
                para.ParagraphFormat.LeftIndent = indentLevel.Value * 36;
            else if (leftIndent.HasValue) para.ParagraphFormat.LeftIndent = leftIndent.Value;

            if (firstLineIndent.HasValue) para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;

            if (tabStops is { Count: > 0 })
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
                    targetPara.ParentNode.InsertBefore(para, targetPara);
                else
                    targetPara.ParentNode.InsertAfter(para, targetPara);
            }
            else
            {
                builder.CurrentParagraph.ParentNode.AppendChild(para);
            }

            // Fix empty paragraphs created after insertion to use Normal style
            // Word automatically creates empty paragraphs after insertion, and they inherit the previous paragraph's style
            // We need to ensure these empty paragraphs use Normal style instead
            // Check all empty paragraphs in the parent node, not just the next sibling
            var parentNode = para.ParentNode;
            if (parentNode != null)
            {
                // Get all paragraphs in the parent node
                var allParagraphs = parentNode.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();

                // Find the index of the inserted paragraph
                var insertedIndex = allParagraphs.IndexOf(para);

                // Check paragraphs after the inserted one
                for (var i = insertedIndex + 1; i < allParagraphs.Count; i++)
                {
                    var nextPara = allParagraphs[i];
                    if (string.IsNullOrWhiteSpace(nextPara.GetText()))
                        // Set empty paragraph to Normal style using StyleIdentifier for more reliable application
                        try
                        {
                            var normalStyle = doc.Styles[StyleIdentifier.Normal];
                            if (normalStyle != null)
                            {
                                // Use StyleIdentifier for more reliable style application
                                nextPara.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                                // Also set Style and StyleName to ensure consistency
                                nextPara.ParagraphFormat.Style = normalStyle;
                                nextPara.ParagraphFormat.StyleName = "Normal";
                                // Clear any direct formatting that might override the style
                                nextPara.ParagraphFormat.ClearFormatting();
                                // Re-apply the style after clearing
                                nextPara.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                                nextPara.ParagraphFormat.Style = normalStyle;
                                nextPara.ParagraphFormat.StyleName = "Normal";
                            }
                        }
                        catch (Exception ex)
                        {
                            // Fallback: try setting StyleName directly
                            Console.Error.WriteLine(
                                $"[WARN] Failed to set paragraph style, trying fallback method: {ex.Message}");
                            try
                            {
                                nextPara.ParagraphFormat.ClearFormatting();
                                nextPara.ParagraphFormat.StyleName = "Normal";
                                nextPara.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                            }
                            catch (Exception ex2)
                            {
                                // If that also fails, skip this paragraph
                                Console.Error.WriteLine(
                                    $"[WARN] Fallback method also failed, skipping paragraph: {ex2.Message}");
                            }
                        }
                    else
                        // Stop at first non-empty paragraph
                        break;
                }
            }

            doc.Save(outputPath);

            var result = "Text added successfully\n";
            if (paragraphIndex.HasValue)
            {
                if (paragraphIndex.Value == -1)
                    result += "Insert position: beginning of document\n";
                else
                    result += $"Insert position: after paragraph #{paragraphIndex.Value}\n";
            }
            else
            {
                result += "Insert position: end of document\n";
            }

            if (!string.IsNullOrEmpty(styleName))
            {
                result += $"Applied style: {styleName}\n";
            }
            else
            {
                result += "Custom formatting:\n";
                if (!string.IsNullOrEmpty(fontNameAscii)) result += $"  Font (ASCII): {fontNameAscii}\n";
                if (!string.IsNullOrEmpty(fontNameFarEast)) result += $"  Font (Far East): {fontNameFarEast}\n";
                if (!string.IsNullOrEmpty(fontName) && string.IsNullOrEmpty(fontNameAscii) &&
                    string.IsNullOrEmpty(fontNameFarEast))
                    result += $"  Font: {fontName}\n";
                if (fontSize.HasValue) result += $"  Font size: {fontSize.Value} pt\n";
                if (bold.HasValue && bold.Value) result += "  Bold\n";
                if (italic.HasValue && italic.Value) result += "  Italic\n";
                if (underline.HasValue && underline.Value) result += "  Underline\n";
                if (!string.IsNullOrEmpty(color)) result += $"  Color: {color}\n";
                if (!string.IsNullOrEmpty(alignment)) result += $"  Alignment: {alignment}\n";
            }

            if (indentLevel.HasValue) result += $"Indent level: {indentLevel.Value} ({indentLevel.Value * 36} pt)\n";
            else if (leftIndent.HasValue) result += $"Left indent: {leftIndent.Value} pt\n";
            if (firstLineIndent.HasValue) result += $"First line indent: {firstLineIndent.Value} pt\n";
            result += $"Output: {outputPath}";

            result += warningMessage;

            return result;
        });
    }
}

/// <summary>
///     Helper class to skip field replacement during text replacement operations
///     Prevents replacement of text inside Word fields (like hyperlinks) unless explicitly requested
/// </summary>
internal class FieldSkipReplacingCallback : IReplacingCallback
{
    /// <summary>
    ///     Determines whether to replace or skip text replacement based on field context
    /// </summary>
    /// <param name="args">Replacing arguments containing match information</param>
    /// <returns>ReplaceAction.Skip if inside a field, ReplaceAction.Replace otherwise</returns>
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Skip replacement if we're inside a field
        if (args.MatchNode.GetAncestor(NodeType.FieldStart) != null ||
            args.MatchNode.GetAncestor(NodeType.FieldSeparator) != null ||
            args.MatchNode.GetAncestor(NodeType.FieldEnd) != null)
            return ReplaceAction.Skip;
        return ReplaceAction.Replace;
    }
}