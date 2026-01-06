using System.ComponentModel;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for text operations in Word documents
///     Merges: WordAddTextTool, WordDeleteTextTool, WordReplaceTextTool, WordSearchTextTool,
///     WordFormatTextTool, WordInsertTextAtPositionTool, WordDeleteTextRangeTool, WordAddTextWithStyleTool
/// </summary>
[McpServerToolType]
public class WordTextTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordTextTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordTextTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word text operation (add, delete, replace, search, format, insert_at_position, delete_range,
    ///     add_with_style).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add, delete, replace, search, format, insert_at_position,
    ///     delete_range, add_with_style.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (for write operations).</param>
    /// <param name="text">Text content (for add, replace, insert_at_position, add_with_style).</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontNameAscii">Font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">Font name for Far East characters.</param>
    /// <param name="fontSize">Font size.</param>
    /// <param name="bold">Bold text.</param>
    /// <param name="italic">Italic text.</param>
    /// <param name="underline">Underline style: none, single, double, dotted, dash.</param>
    /// <param name="color">Text color hex or name.</param>
    /// <param name="strikethrough">Strikethrough text.</param>
    /// <param name="superscript">Superscript text.</param>
    /// <param name="subscript">Subscript text.</param>
    /// <param name="find">Text to find (for replace).</param>
    /// <param name="replace">Replacement text (for replace).</param>
    /// <param name="searchText">Text to search for (for search/delete).</param>
    /// <param name="caseSensitive">Case sensitive search (for search/replace).</param>
    /// <param name="useRegex">Use regular expression (for search/replace).</param>
    /// <param name="replaceInFields">Replace text inside fields (for replace).</param>
    /// <param name="maxResults">Maximum number of results to return (for search).</param>
    /// <param name="contextLength">Number of characters to show for context (for search).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based).</param>
    /// <param name="runIndex">Run index within paragraph (0-based).</param>
    /// <param name="startParagraphIndex">Start paragraph index (for delete_range).</param>
    /// <param name="startRunIndex">Start run index (for delete_range).</param>
    /// <param name="endParagraphIndex">End paragraph index (for delete_range).</param>
    /// <param name="endRunIndex">End run index (for delete_range).</param>
    /// <param name="insertParagraphIndex">Paragraph index for insert_at_position.</param>
    /// <param name="charIndex">Character index for insert_at_position.</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="insertBefore">Insert before position (for insert_at_position).</param>
    /// <param name="startCharIndex">Start character index (for delete_range).</param>
    /// <param name="endCharIndex">End character index (for delete_range).</param>
    /// <param name="styleName">Style name (for add_with_style).</param>
    /// <param name="alignment">Text alignment (for add_with_style).</param>
    /// <param name="indentLevel">Indentation level (for add_with_style).</param>
    /// <param name="leftIndent">Left indentation in points (for add_with_style).</param>
    /// <param name="firstLineIndent">First line indentation in points (for add_with_style).</param>
    /// <param name="paragraphIndexForAdd">Paragraph index to insert after (for add_with_style).</param>
    /// <param name="tabStops">Custom tab stops as JSON array (for add_with_style).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for search operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_text")]
    [Description(
        @"Perform text operations in Word documents. Supports 8 operations: add, delete, replace, search, format, insert_at_position, delete_range, add_with_style.

Usage examples:
- Add text: word_text(operation='add', path='doc.docx', text='Hello World')
- Add formatted text: word_text(operation='add', path='doc.docx', text='Bold text', bold=true)
- Replace text: word_text(operation='replace', path='doc.docx', find='old', replace='new')
- Search text: word_text(operation='search', path='doc.docx', searchText='keyword')
- Format text: word_text(operation='format', path='doc.docx', paragraphIndex=0, runIndex=0, bold=true)
- Insert at position: word_text(operation='insert_at_position', path='doc.docx', paragraphIndex=0, runIndex=0, text='Inserted')
- Delete text: word_text(operation='delete', path='doc.docx', searchText='text to delete') or word_text(operation='delete', path='doc.docx', startParagraphIndex=0, endParagraphIndex=0)
- Delete range: word_text(operation='delete_range', path='doc.docx', startParagraphIndex=0, startRunIndex=0, endParagraphIndex=0, endRunIndex=5)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add text at document end (required params: path, text)
- 'delete': Delete text (required params: path, searchText OR startParagraphIndex+endParagraphIndex)
- 'replace': Replace text (required params: path, find, replace)
- 'search': Search for text (required params: path, searchText)
- 'format': Format existing text (required params: path, paragraphIndex, runIndex)
- 'insert_at_position': Insert text at specific position (required params: path, paragraphIndex, runIndex, text)
- 'delete_range': Delete text range (required params: path, startParagraphIndex, startRunIndex, endParagraphIndex, endRunIndex)
- 'add_with_style': Add text with style (required params: path, text, styleName)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (if not provided, overwrites input, for write operations)")]
        string? outputPath = null,
        // Common parameters
        [Description("Text content (required for add, replace, insert_at_position, add_with_style operations)")]
        string? text = null,
        // Add/AddWithStyle parameters
        [Description("Font name (optional, e.g., 'Arial')")]
        string? fontName = null,
        [Description("Font name for ASCII characters (English, e.g., 'Times New Roman')")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (Chinese/Japanese/Korean, optional)")]
        string? fontNameFarEast = null,
        [Description("Font size (optional)")] double? fontSize = null,
        [Description("Bold text (optional)")] bool? bold = null,
        [Description("Italic text (optional)")]
        bool? italic = null,
        [Description("Underline style: none, single, double, dotted, dash (optional, for add/format operations)")]
        string? underline = null,
        [Description(
            "Text color (hex format like 'FF0000' or '#FF0000' for red, or name like 'Red', 'Blue', optional)")]
        string? color = null,
        [Description("Strikethrough (optional, for add/format operations)")]
        bool? strikethrough = null,
        [Description("Superscript (optional, for add/format operations)")]
        bool? superscript = null,
        [Description("Subscript (optional, for add/format operations)")]
        bool? subscript = null,
        // Replace parameters
        [Description("Text to find (required for replace operation)")]
        string? find = null,
        [Description("Replacement text (required for replace operation)")]
        string? replace = null,
        [Description("Use regex matching (optional, for replace/search operations)")]
        bool useRegex = false,
        [Description(
            "Replace text inside fields (optional, default: false). If false, fields like hyperlinks will be excluded from replacement to preserve their functionality")]
        bool replaceInFields = false,
        // Search parameters
        [Description(
            "Text to search for (required for search operation, optional for delete operation as alternative to startParagraphIndex+endParagraphIndex)")]
        string? searchText = null,
        [Description("Case sensitive search (optional, default: false, for search operation)")]
        bool caseSensitive = false,
        [Description("Maximum number of results to return (optional, default: 50, for search operation)")]
        int maxResults = 50,
        [Description(
            "Number of characters to show before and after match for context (optional, default: 50, for search operation)")]
        int contextLength = 50,
        // Delete parameters
        [Description("Start paragraph index (0-based, required for delete operation if searchText is not provided)")]
        int? startParagraphIndex = null,
        [Description("Start run index within start paragraph (0-based, optional, default: 0, for delete operation)")]
        int startRunIndex = 0,
        [Description(
            "End paragraph index (0-based, inclusive, required for delete operation if searchText is not provided)")]
        int? endParagraphIndex = null,
        [Description(
            "End run index within end paragraph (0-based, inclusive, optional, default: last run, for delete operation)")]
        int? endRunIndex = null,
        // Format parameters
        [Description("Paragraph index (0-based, required for format operation)")]
        int? paragraphIndex = null,
        [Description(
            "Run index within the paragraph (0-based, optional, formats all runs if not provided, for format operation)")]
        int? runIndex = null,
        // Insert at position parameters
        [Description("Paragraph index (0-based, required for insert_at_position operation)")]
        int? insertParagraphIndex = null,
        [Description("Character index within paragraph (0-based, required for insert_at_position operation)")]
        int? charIndex = null,
        [Description(
            "Section index (0-based, optional, default: 0, for format/insert_at_position/delete_range operations)")]
        int? sectionIndex = null,
        [Description(
            "Insert before position (optional, default: false, inserts after, for insert_at_position operation)")]
        bool insertBefore = false,
        // Delete range parameters
        [Description("Start character index within paragraph (0-based, required for delete_range operation)")]
        int? startCharIndex = null,
        [Description("End character index within paragraph (0-based, required for delete_range operation)")]
        int? endCharIndex = null,
        // AddWithStyle parameters
        [Description("Style name to apply (e.g., 'Heading 1', 'Normal', optional, for add_with_style operation)")]
        string? styleName = null,
        [Description("Text alignment: left, center, right, justify (optional, for add_with_style operation)")]
        string? alignment = null,
        [Description(
            "Indentation level (0-8, where each level = 36 points / 0.5 inch, optional, for add_with_style operation)")]
        int? indentLevel = null,
        [Description("Left indentation in points (optional, for add_with_style operation)")]
        double? leftIndent = null,
        [Description(
            "First line indentation in points (positive = indent first line, negative = hanging indent, optional, for add_with_style operation)")]
        double? firstLineIndent = null,
        [Description(
            "Index of the paragraph to insert after (0-based, optional, for add_with_style operation). Use -1 to insert at the beginning.")]
        int? paragraphIndexForAdd = null,
        [Description(
            "Custom tab stops as JSON array (optional, for add_with_style operation). Example: [{\"position\":72,\"alignment\":\"Left\",\"leader\":\"None\"}]")]
        string? tabStops = null)
    {
        var effectiveOutputPath = outputPath ?? path;
        if (!string.IsNullOrEmpty(effectiveOutputPath))
            SecurityHelper.ValidateFilePath(effectiveOutputPath, "outputPath", true);

        // Parse tabStops from JSON string
        JsonArray? tabStopsArray = null;
        if (!string.IsNullOrEmpty(tabStops))
            tabStopsArray = JsonNode.Parse(tabStops) as JsonArray;

        return operation.ToLower() switch
        {
            "add" => AddText(path, sessionId, effectiveOutputPath, text, fontName, fontSize, bold, italic, underline,
                color, strikethrough, superscript, subscript),
            "delete" => DeleteText(path, sessionId, effectiveOutputPath, searchText, startParagraphIndex, startRunIndex,
                endParagraphIndex, endRunIndex),
            "replace" => ReplaceText(path, sessionId, effectiveOutputPath, find, replace, useRegex, replaceInFields),
            "search" => SearchText(path, sessionId, searchText, useRegex, caseSensitive, maxResults, contextLength),
            "format" => FormatText(path, sessionId, effectiveOutputPath, paragraphIndex, runIndex, sectionIndex ?? 0,
                fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline, color, strikethrough,
                superscript, subscript),
            "insert_at_position" => InsertTextAtPosition(path, sessionId, effectiveOutputPath, insertParagraphIndex,
                charIndex, sectionIndex, text, insertBefore),
            "delete_range" => DeleteTextRange(path, sessionId, effectiveOutputPath, startParagraphIndex ?? 0,
                startCharIndex ?? 0, endParagraphIndex ?? 0, endCharIndex ?? 0, sectionIndex),
            "add_with_style" => AddTextWithStyle(path, sessionId, effectiveOutputPath, text, styleName, fontName,
                fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline != null, color, alignment,
                indentLevel, leftIndent, firstLineIndent, tabStopsArray, paragraphIndexForAdd),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds text to the document.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The text content to add.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text should be bold.</param>
    /// <param name="italic">Whether the text should be italic.</param>
    /// <param name="underline">The underline style.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="strikethrough">Whether the text should have strikethrough.</param>
    /// <param name="superscript">Whether the text should be superscript.</param>
    /// <param name="subscript">Whether the text should be subscript.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is null or empty.</exception>
    private string AddText(string? path, string? sessionId, string? outputPath, string? text,
        string? fontName, double? fontSize, bool? bold, bool? italic, string? underline, string? color,
        bool? strikethrough, bool? superscript, bool? subscript)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var doc = ctx.Document;

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

            FontHelper.Word.ApplyFontSettings(
                builder,
                fontName,
                fontSize: fontSize,
                bold: bold,
                italic: italic,
                underline: underline,
                color: color,
                strikethrough: strikethrough,
                superscript: superscript,
                subscript: subscript
            );

            var currentPara = builder.CurrentParagraph;
            var runsBefore = 0;
            if (currentPara != null) runsBefore = currentPara.GetChildNodes(NodeType.Run, false).Count;

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
                                bold: bold,
                                italic: italic,
                                underline: underline,
                                strikethrough: strikethrough,
                                superscript: superscript,
                                subscript: subscript
                            );
                        }
                    }
            }
        }

        ctx.Save(outputPath);

        List<string> formatInfo = [];
        if (bold == true) formatInfo.Add("bold");
        if (italic == true) formatInfo.Add("italic");
        if (!string.IsNullOrEmpty(underline) && underline != "none") formatInfo.Add($"underline({underline})");
        if (strikethrough == true) formatInfo.Add("strikethrough");
        if (superscript == true) formatInfo.Add("superscript");
        if (subscript == true) formatInfo.Add("subscript");

        var result = "Text added to document successfully\n";
        if (formatInfo.Count > 0) result += $"Applied formats: {string.Join(", ", formatInfo)}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Deletes text from the document.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="searchText">The text to search for and delete.</param>
    /// <param name="startParagraphIdx">The starting paragraph index.</param>
    /// <param name="startRunIdx">The starting run index within the paragraph.</param>
    /// <param name="endParagraphIdx">The ending paragraph index.</param>
    /// <param name="endRunIdx">The ending run index within the paragraph.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the specified paragraph cannot be found.</exception>
    private string DeleteText(string? path, string? sessionId, string? outputPath,
        string? searchText, int? startParagraphIdx, int startRunIdx, int? endParagraphIdx, int? endRunIdx)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        var startParagraphIndex = startParagraphIdx;
        var startRunIndex = startRunIdx;
        var endParagraphIndex = endParagraphIdx;
        var endRunIndex = endRunIdx;

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
                    var startRunIdx2 = 0;
                    var endRunIdx2 = runs.Count - 1;

                    for (var r = 0; r < runs.Count; r++)
                    {
                        if (runs[r] is not Run run) continue;

                        var runLength = run.Text.Length;
                        if (charCount + runLength > textIndex)
                        {
                            startRunIdx2 = r;
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
                            endRunIdx2 = r;
                            break;
                        }

                        charCount += runLength;
                    }

                    startParagraphIndex = p;
                    endParagraphIndex = p;
                    startRunIndex = startRunIdx2;
                    endRunIndex = endRunIdx2;
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

        ctx.Save(outputPath);

        var preview = deletedText.Length > 50 ? deletedText.Substring(0, 50) + "..." : deletedText;

        var result = "Text deleted successfully\n";
        if (!string.IsNullOrEmpty(searchText)) result += $"Deleted text: {searchText}\n";
        result +=
            $"Range: Paragraph {startParagraphIndex.Value} Run {startRunIndex} to Paragraph {endParagraphIndex.Value} Run {endRunIndex ?? -1}\n";
        if (!string.IsNullOrEmpty(preview)) result += $"Deleted content preview: {preview}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Replaces text in the document.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="find">The text to find.</param>
    /// <param name="replace">The replacement text.</param>
    /// <param name="useRegex">Whether to use regex matching.</param>
    /// <param name="replaceInFields">Whether to replace text inside fields like hyperlinks.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when find or replace parameters are missing.</exception>
    private string ReplaceText(string? path, string? sessionId, string? outputPath,
        string? find, string? replace, bool useRegex, bool replaceInFields)
    {
        if (string.IsNullOrEmpty(find))
            throw new ArgumentException("find is required for replace operation");
        if (replace == null)
            throw new ArgumentException("replace is required for replace operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var doc = ctx.Document;

        var options = new FindReplaceOptions();

        // Fields (like hyperlinks) should not be replaced unless explicitly requested
        if (!replaceInFields) options.ReplacingCallback = new FieldSkipReplacingCallback();

        if (useRegex)
            doc.Range.Replace(new Regex(find), replace, options);
        else
            doc.Range.Replace(find, replace, options);

        ctx.Save(outputPath);

        var result = $"Text replaced in document.\n{ctx.GetOutputMessage(outputPath)}";
        if (!replaceInFields)
            result +=
                "\nNote: Fields (such as hyperlinks) were excluded from replacement to preserve their functionality.";
        return result;
    }

    /// <summary>
    ///     Searches for text in the document.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="searchText">The text to search for.</param>
    /// <param name="useRegex">Whether to use regex matching.</param>
    /// <param name="caseSensitive">Whether the search is case sensitive.</param>
    /// <param name="maxResults">The maximum number of results to return.</param>
    /// <param name="contextLength">The number of characters to show before and after each match.</param>
    /// <returns>A formatted string containing the search results.</returns>
    /// <exception cref="ArgumentException">Thrown when searchText is null or empty.</exception>
    private string SearchText(string? path, string? sessionId,
        string? searchText, bool useRegex, bool caseSensitive, int maxResults, int contextLength)
    {
        if (string.IsNullOrEmpty(searchText))
            throw new ArgumentException("searchText is required for search operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var doc = ctx.Document;
        var result = new StringBuilder();
        List<(string text, int paragraphIndex, string context)> matches = [];

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
    }

    /// <summary>
    ///     Extracts context around a matched text for search results display
    /// </summary>
    /// <param name="text">The full paragraph text</param>
    /// <param name="matchIndex">Starting index of the match</param>
    /// <param name="matchLength">Length of the matched text</param>
    /// <param name="contextLength">Number of characters to include before and after the match</param>
    /// <returns>Context string with the match highlighted using brackets</returns>
    private static string GetContext(string text, int matchIndex, int matchLength, int contextLength)
    {
        var start = Math.Max(0, matchIndex - contextLength);
        var end = Math.Min(text.Length, matchIndex + matchLength + contextLength);

        var context = text.Substring(start, end - start);

        context = context.Replace("\r", "").Replace("\n", " ").Trim();

        var highlightStart = matchIndex - start;
        var highlightEnd = highlightStart + matchLength;

        if (highlightStart >= 0 && highlightEnd <= context.Length)
            context = context.Substring(0, highlightStart) +
                      "[" + context.Substring(highlightStart, matchLength) + "]" +
                      context.Substring(highlightEnd);

        return context;
    }

    /// <summary>
    ///     Formats text in the document.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="runIndex">The run index within the paragraph (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="fontName">The font name to apply.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text should be bold.</param>
    /// <param name="italic">Whether the text should be italic.</param>
    /// <param name="underline">The underline style.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="strikethrough">Whether the text should have strikethrough.</param>
    /// <param name="superscript">Whether the text should be superscript.</param>
    /// <param name="subscript">Whether the text should be subscript.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraphIndex is missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the paragraph has no runs and cannot create new ones.</exception>
    private string FormatText(string? path, string? sessionId, string? outputPath,
        int? paragraphIndex, int? runIndex, int sectionIndex,
        string? fontName, string? fontNameAscii, string? fontNameFarEast,
        double? fontSize, bool? bold, bool? italic, string? underline, string? color,
        bool? strikethrough, bool? superscript, bool? subscript)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex is required for format operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var doc = ctx.Document;

        // Use section-based paragraph indexing to match GetRunFormat behavior
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException(
                $"sectionIndex {sectionIndex} is out of range (document has {doc.Sections.Count} sections)");

        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (section {sectionIndex} body has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paragraphIndex.Value];

        var runs = para.GetChildNodes(NodeType.Run, false);
        if (runs == null || runs.Count == 0)
        {
            var newRun = new Run(doc, "");
            para.AppendChild(newRun);
            runs = para.GetChildNodes(NodeType.Run, false);
            if (runs == null || runs.Count == 0)
                throw new InvalidOperationException(
                    $"Paragraph #{paragraphIndex.Value} has no Run nodes and cannot create new Run node");
        }

        List<string> changes = [];
        List<Run> runsToFormat = [];

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

        ctx.Save(outputPath);

        var result = "Run-level formatting set successfully\n";
        result += $"Paragraph index: {paragraphIndex.Value}\n";
        if (runIndex.HasValue)
            result += $"Run index: {runIndex.Value}\n";
        else
            result += $"Formatted Runs: {runsToFormat.Count}\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes.Distinct())}\n";
        else
            result += "No change parameters provided\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Inserts text at a specific position.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="charIndex">The character index within the paragraph (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="text">The text to insert.</param>
    /// <param name="insertBefore">Whether to insert before the position.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    private string InsertTextAtPosition(string? path, string? sessionId, string? outputPath,
        int? paragraphIndex, int? charIndex, int? sectionIndex, string? text, bool insertBefore)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("insertParagraphIndex is required for insert_at_position operation");
        if (!charIndex.HasValue)
            throw new ArgumentException("charIndex is required for insert_at_position operation");
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for insert_at_position operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var doc = ctx.Document;
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var section = doc.Sections[sectionIdx];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");

        var para = paragraphs[paragraphIndex.Value];
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var totalChars = 0;
        var targetRunIndex = -1;
        var targetRunCharIndex = 0;

        for (var i = 0; i < runs.Count; i++)
        {
            var runLength = runs[i].Text.Length;
            if (totalChars + runLength >= charIndex.Value)
            {
                targetRunIndex = i;
                targetRunCharIndex = charIndex.Value - totalChars;
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
                builder.MoveToParagraph(paragraphIndex.Value, para.GetText().Length);
                builder.Write(text);
            }
        }
        else
        {
            var targetRun = runs[targetRunIndex];
            targetRun.Text = targetRun.Text.Insert(targetRunCharIndex, text);
        }

        ctx.Save(outputPath);
        return $"Text inserted at position.\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes text in a range.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="startParagraphIndex">The starting paragraph index (0-based).</param>
    /// <param name="startCharIndex">The starting character index within the paragraph.</param>
    /// <param name="endParagraphIndex">The ending paragraph index (0-based).</param>
    /// <param name="endCharIndex">The ending character index within the paragraph.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range.</exception>
    private string DeleteTextRange(string? path, string? sessionId, string? outputPath,
        int startParagraphIndex, int startCharIndex, int endParagraphIndex, int endCharIndex, int? sectionIndex)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var doc = ctx.Document;
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

        ctx.Save(outputPath);
        return $"Text range deleted.\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Adds text with a specific style.
    /// </summary>
    /// <param name="path">The document file path.</param>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The text content to add.</param>
    /// <param name="styleName">The style name to apply.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text should be bold.</param>
    /// <param name="italic">Whether the text should be italic.</param>
    /// <param name="underline">Whether the text should be underlined.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="alignment">The paragraph alignment.</param>
    /// <param name="indentLevel">The indentation level (0-8).</param>
    /// <param name="leftIndent">The left indentation in points.</param>
    /// <param name="firstLineIndent">The first line indentation in points.</param>
    /// <param name="tabStops">Custom tab stops as JSON array.</param>
    /// <param name="paragraphIndex">The paragraph index to insert after.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is null or empty or style is not found.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the style cannot be applied.</exception>
    private string AddTextWithStyle(string? path, string? sessionId, string? outputPath,
        string? text, string? styleName, string? fontName, string? fontNameAscii, string? fontNameFarEast,
        double? fontSize, bool? bold, bool? italic, bool? underline, string? color,
        string? alignment, int? indentLevel, double? leftIndent, double? firstLineIndent,
        JsonArray? tabStops, int? paragraphIndex)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add_with_style operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);
        var doc = ctx.Document;
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
                "\nNote: When using both styleName and custom parameters, custom parameters will override corresponding properties in the style.\n" +
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

        ctx.Save(outputPath);

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
        result += ctx.GetOutputMessage(outputPath);

        result += warningMessage;

        return result;
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