using System.ComponentModel;
using System.Drawing;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for paragraph operations in Word documents
///     Merges: WordInsertParagraphTool, WordDeleteParagraphTool, WordEditParagraphTool,
///     WordGetParagraphsTool, WordGetParagraphFormatTool, WordCopyParagraphFormatTool, WordMergeParagraphsTool
/// </summary>
[McpServerToolType]
public class WordParagraphTool
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
    ///     Initializes a new instance of the WordParagraphTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordParagraphTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word paragraph operation (insert, delete, edit, get, get_format, copy_format, merge).
    /// </summary>
    /// <param name="operation">The operation to perform: insert, delete, edit, get, get_format, copy_format, merge.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, -1 for last paragraph).</param>
    /// <param name="text">Text content for the paragraph.</param>
    /// <param name="styleName">Style name to apply (e.g., 'Heading 1', 'Normal').</param>
    /// <param name="alignment">Text alignment: left, center, right, justify.</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="includeEmpty">Include empty paragraphs (default: true).</param>
    /// <param name="styleFilter">Filter by style name.</param>
    /// <param name="includeCommentParagraphs">Include paragraphs inside nested structures (default: true).</param>
    /// <param name="includeTextboxParagraphs">Include paragraphs inside TextBox/Shape objects (default: true).</param>
    /// <param name="includeRunDetails">Include detailed run-level formatting (default: true).</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontNameAscii">Font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">Font name for Far East characters.</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="bold">Bold text.</param>
    /// <param name="italic">Italic text.</param>
    /// <param name="underline">Underline text.</param>
    /// <param name="color">Text color hex.</param>
    /// <param name="indentLeft">Left indent in points.</param>
    /// <param name="indentRight">Right indent in points.</param>
    /// <param name="firstLineIndent">First line indent in points.</param>
    /// <param name="spaceBefore">Space before paragraph in points.</param>
    /// <param name="spaceAfter">Space after paragraph in points.</param>
    /// <param name="lineSpacing">Line spacing multiplier.</param>
    /// <param name="lineSpacingRule">Line spacing rule: single, oneAndHalf, double, atLeast, exactly, multiple.</param>
    /// <param name="tabStops">Custom tab stops array.</param>
    /// <param name="sourceParagraphIndex">Source paragraph index (for copy_format).</param>
    /// <param name="targetParagraphIndex">Target paragraph index (for copy_format).</param>
    /// <param name="startParagraphIndex">Start paragraph index (for merge).</param>
    /// <param name="endParagraphIndex">End paragraph index (for merge).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_paragraph")]
    [Description(
        @"Manage paragraphs in Word documents. Supports 7 operations: insert, delete, edit, get, get_format, copy_format, merge.

Usage examples:
- Insert paragraph: word_paragraph(operation='insert', path='doc.docx', paragraphIndex=0, text='New paragraph')
- Delete paragraph: word_paragraph(operation='delete', path='doc.docx', paragraphIndex=0)
- Edit format: word_paragraph(operation='edit', path='doc.docx', paragraphIndex=0, alignment='center', fontSize=14)
- Get paragraph: word_paragraph(operation='get', path='doc.docx', paragraphIndex=0)
- Get format: word_paragraph(operation='get_format', path='doc.docx', paragraphIndex=0)
- Copy format: word_paragraph(operation='copy_format', path='doc.docx', sourceParagraphIndex=0, targetParagraphIndex=1)
- Merge paragraphs: word_paragraph(operation='merge', path='doc.docx', startParagraphIndex=0, endParagraphIndex=2)

Important notes for 'get' operation:
- By default, returns ALL paragraphs in the document structure, including paragraphs inside Comment objects, table cells, and TextBoxes
- Use includeCommentParagraphs=false to get only Body paragraphs (visible in document body, excluding table cells and comments)
- Each paragraph shows its ParentNode type to help identify its location
- Paragraphs inside Comment objects are marked with '[Comment]' in the location field
- Paragraphs inside table cells are marked with '[Cell]' in the location field
- Paragraphs inside TextBoxes are marked with '[TextBox]' in the location field
- To check paragraph styles in table cells, use includeCommentParagraphs=true")]
    public string Execute(
        [Description("Operation: insert, delete, edit, get, get_format, copy_format, merge")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Paragraph index (0-based, -1 for last paragraph)")]
        int? paragraphIndex = null,
        [Description("Text content for the paragraph")]
        string? text = null,
        [Description("Style name to apply (e.g., 'Heading 1', 'Normal')")]
        string? styleName = null,
        [Description("Text alignment: left, center, right, justify")]
        string? alignment = null,
        [Description("Section index (0-based)")]
        int? sectionIndex = null,
        [Description("Include empty paragraphs (default: true)")]
        bool includeEmpty = true,
        [Description("Filter by style name")] string? styleFilter = null,
        [Description("Include paragraphs inside nested structures (default: true)")]
        bool includeCommentParagraphs = true,
        [Description("Include paragraphs inside TextBox/Shape objects (default: true)")]
        bool includeTextboxParagraphs = true,
        [Description("Include detailed run-level formatting (default: true)")]
        bool includeRunDetails = true,
        [Description("Font name")] string? fontName = null,
        [Description("Font name for ASCII characters")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters")]
        string? fontNameFarEast = null,
        [Description("Font size in points")] double? fontSize = null,
        [Description("Bold text")] bool? bold = null,
        [Description("Italic text")] bool? italic = null,
        [Description("Underline text")] bool? underline = null,
        [Description("Font color (hex format, e.g., '000000')")]
        string? color = null,
        [Description("Left indent in points")] double? indentLeft = null,
        [Description("Right indent in points")]
        double? indentRight = null,
        [Description("First line indent in points")]
        double? firstLineIndent = null,
        [Description("Space before paragraph in points")]
        double? spaceBefore = null,
        [Description("Space after paragraph in points")]
        double? spaceAfter = null,
        [Description("Line spacing value")] double? lineSpacing = null,
        [Description("Line spacing rule: single, oneAndHalf, double, atLeast, exactly, multiple")]
        string? lineSpacingRule = null,
        [Description("Custom tab stops array")]
        JsonArray? tabStops = null,
        [Description("Source paragraph index (for copy_format)")]
        int? sourceParagraphIndex = null,
        [Description("Target paragraph index (for copy_format)")]
        int? targetParagraphIndex = null,
        [Description("Start paragraph index (for merge)")]
        int? startParagraphIndex = null,
        [Description("End paragraph index (for merge)")]
        int? endParagraphIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "insert" => InsertParagraph(ctx, outputPath, text, paragraphIndex, styleName, alignment, indentLeft,
                indentRight, firstLineIndent, spaceBefore, spaceAfter),
            "delete" => DeleteParagraph(ctx, outputPath, paragraphIndex),
            "edit" => EditParagraph(ctx, outputPath, paragraphIndex, sectionIndex, text, styleName, alignment, fontName,
                fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline, color, indentLeft, indentRight,
                firstLineIndent, spaceBefore, spaceAfter, lineSpacing, lineSpacingRule, tabStops),
            "get" => GetParagraphs(ctx, sectionIndex, includeEmpty, styleFilter, includeCommentParagraphs,
                includeTextboxParagraphs),
            "get_format" => GetParagraphFormat(ctx, paragraphIndex, includeRunDetails),
            "copy_format" => CopyParagraphFormat(ctx, outputPath, sourceParagraphIndex, targetParagraphIndex),
            "merge" => MergeParagraphs(ctx, outputPath, startParagraphIndex, endParagraphIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a new paragraph at the specified position with optional formatting.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The text content for the paragraph.</param>
    /// <param name="paragraphIndex">The paragraph index to insert at (-1 for beginning).</param>
    /// <param name="styleName">The style name to apply.</param>
    /// <param name="alignment">The text alignment.</param>
    /// <param name="indentLeft">The left indentation in points.</param>
    /// <param name="indentRight">The right indentation in points.</param>
    /// <param name="firstLineIndent">The first line indentation in points.</param>
    /// <param name="spaceBefore">The space before paragraph in points.</param>
    /// <param name="spaceAfter">The space after paragraph in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is null or empty, index is out of range, or style is not found.</exception>
    private static string InsertParagraph(DocumentContext<Document> ctx, string? outputPath, string? text,
        int? paragraphIndex, string? styleName, string? alignment, double? indentLeft, double? indentRight,
        double? firstLineIndent, double? spaceBefore, double? spaceAfter)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for insert operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        Paragraph? targetPara = null;
        var insertPosition = "end of document";

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                {
                    targetPara = paragraphs[0] as Paragraph;
                    insertPosition = "beginning of document";
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                insertPosition = $"after paragraph #{paragraphIndex.Value}";
            }
            else
            {
                var validRange = paragraphs.Count > 0 ? $"0-{paragraphs.Count - 1}" : "none (document is empty)";
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: {validRange}, or -1 for beginning).");
            }
        }

        var para = new Paragraph(doc);
        var run = new Run(doc, text);
        para.AppendChild(run);

        if (!string.IsNullOrEmpty(styleName))
        {
            var style = doc.Styles[styleName];
            if (style != null)
                para.ParagraphFormat.StyleName = styleName;
            else
                throw new ArgumentException(
                    $"Style '{styleName}' not found. Use word_get_styles tool to view available styles");
        }

        if (!string.IsNullOrEmpty(alignment))
            para.ParagraphFormat.Alignment = GetAlignment(alignment);

        if (indentLeft.HasValue) para.ParagraphFormat.LeftIndent = indentLeft.Value;
        if (indentRight.HasValue) para.ParagraphFormat.RightIndent = indentRight.Value;
        if (firstLineIndent.HasValue) para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
        if (spaceBefore.HasValue) para.ParagraphFormat.SpaceBefore = spaceBefore.Value;
        if (spaceAfter.HasValue) para.ParagraphFormat.SpaceAfter = spaceAfter.Value;

        if (targetPara != null)
        {
            if (paragraphIndex!.Value == -1)
                targetPara.ParentNode.InsertBefore(para, targetPara);
            else
                targetPara.ParentNode.InsertAfter(para, targetPara);
        }
        else
        {
            var body = doc.FirstSection.Body;
            body.AppendChild(para);
        }

        ctx.Save(outputPath);

        var result = "Paragraph inserted successfully\n";
        result += $"Insert position: {insertPosition}\n";
        if (!string.IsNullOrEmpty(styleName)) result += $"Applied style: {styleName}\n";
        if (!string.IsNullOrEmpty(alignment)) result += $"Alignment: {alignment}\n";
        result += $"Document paragraph count: {paragraphs.Count + 1}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Deletes a paragraph at the specified index from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The paragraph index to delete (-1 for last).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraphIndex is null or index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the paragraph cannot be found.</exception>
    private static string DeleteParagraph(DocumentContext<Document> ctx, string? outputPath, int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for delete operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        var idx = paragraphIndex.Value;
        if (idx == -1)
        {
            if (paragraphs.Count == 0)
                throw new ArgumentException("Cannot delete paragraph: document has no paragraphs");
            idx = paragraphs.Count - 1;
        }

        if (idx < 0 || idx >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {idx} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}, or -1 for last).");

        var paragraphToDelete = paragraphs[idx] as Paragraph;
        if (paragraphToDelete == null)
            throw new InvalidOperationException($"Unable to get paragraph at index {idx}");

        var textPreview = paragraphToDelete.GetText().Trim();
        if (textPreview.Length > 50) textPreview = textPreview.Substring(0, 50) + "...";

        paragraphToDelete.Remove();

        ctx.Save(outputPath);

        var result = $"Paragraph #{idx} deleted successfully\n";
        if (!string.IsNullOrEmpty(textPreview)) result += $"Content preview: {textPreview}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Edits paragraph content and formatting at the specified index.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The paragraph index to edit (-1 for last).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="text">The new text content.</param>
    /// <param name="styleName">The style name to apply.</param>
    /// <param name="alignment">The text alignment.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text should be bold.</param>
    /// <param name="italic">Whether the text should be italic.</param>
    /// <param name="underline">Whether the text should be underlined.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="indentLeft">The left indentation in points.</param>
    /// <param name="indentRight">The right indentation in points.</param>
    /// <param name="firstLineIndent">The first line indentation in points.</param>
    /// <param name="spaceBefore">The space before paragraph in points.</param>
    /// <param name="spaceAfter">The space after paragraph in points.</param>
    /// <param name="lineSpacing">The line spacing value.</param>
    /// <param name="lineSpacingRule">The line spacing rule.</param>
    /// <param name="tabStops">Custom tab stops array.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when paragraphIndex is null, indices are out of range, or style is not
    ///     found.
    /// </exception>
    private static string EditParagraph(DocumentContext<Document> ctx, string? outputPath, int? paragraphIndex,
        int? sectionIndex, string? text, string? styleName, string? alignment, string? fontName, string? fontNameAscii,
        string? fontNameFarEast, double? fontSize, bool? bold, bool? italic, bool? underline, string? color,
        double? indentLeft, double? indentRight, double? firstLineIndent, double? spaceBefore, double? spaceAfter,
        double? lineSpacing, string? lineSpacingRule, JsonArray? tabStops)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for edit operation");

        var doc = ctx.Document;
        var secIdx = sectionIndex ?? 0;

        // Handle paragraphIndex=-1 (document end)
        if (paragraphIndex.Value == -1)
        {
            var lastSection = doc.LastSection;
            var bodyParagraphs = lastSection.Body.GetChildNodes(NodeType.Paragraph, false);
            if (bodyParagraphs.Count > 0)
            {
                paragraphIndex = bodyParagraphs.Count - 1;
                secIdx = doc.Sections.Count - 1;
            }
            else
            {
                throw new ArgumentException(
                    "Cannot edit paragraph: document has no paragraphs. Use insert operation to add paragraphs first.");
            }
        }

        if (secIdx < 0 || secIdx >= doc.Sections.Count)
            throw new ArgumentException(
                $"Section index {secIdx} out of range (total sections: {doc.Sections.Count}, valid range: 0-{doc.Sections.Count - 1})");

        var section = doc.Sections[secIdx];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} out of range (total paragraphs: {paragraphs.Count}, valid range: 0-{paragraphs.Count - 1})");

        var para = paragraphs[paragraphIndex.Value];
        var builder = new DocumentBuilder(doc);
        builder.MoveTo(para.FirstChild ?? para);

        var underlineStr = underline.HasValue ? underline.Value ? "single" : "none" : null;

        FontHelper.Word.ApplyFontSettings(builder, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic,
            underlineStr, color);

        var paraFormat = para.ParagraphFormat;

        if (!string.IsNullOrEmpty(alignment))
            paraFormat.Alignment = GetAlignment(alignment);

        if (indentLeft.HasValue) paraFormat.LeftIndent = indentLeft.Value;
        if (indentRight.HasValue) paraFormat.RightIndent = indentRight.Value;
        if (firstLineIndent.HasValue) paraFormat.FirstLineIndent = firstLineIndent.Value;
        if (spaceBefore.HasValue) paraFormat.SpaceBefore = spaceBefore.Value;
        if (spaceAfter.HasValue) paraFormat.SpaceAfter = spaceAfter.Value;

        if (lineSpacing.HasValue || !string.IsNullOrEmpty(lineSpacingRule))
        {
            var rule = GetLineSpacingRule(lineSpacingRule ?? "single");
            paraFormat.LineSpacingRule = rule;

            if (lineSpacing.HasValue)
                paraFormat.LineSpacing = lineSpacing.Value;
            else
                paraFormat.LineSpacing = (lineSpacingRule ?? "single").ToLower() switch
                {
                    "single" => 1.0,
                    "oneandhalf" => 1.5,
                    "double" => 2.0,
                    _ => 1.0
                };
        }

        if (!string.IsNullOrEmpty(styleName))
        {
            var style = doc.Styles[styleName];
            if (style != null)
            {
                var isEmpty = string.IsNullOrWhiteSpace(para.GetText());
                if (isEmpty) paraFormat.ClearFormatting();
                paraFormat.Style = style;
                paraFormat.StyleName = styleName;
            }
            else
            {
                throw new ArgumentException(
                    $"Style '{styleName}' not found. Use word_get_styles tool to view available styles");
            }
        }

        if (tabStops is { Count: > 0 })
        {
            paraFormat.TabStops.Clear();
            foreach (var ts in tabStops)
            {
                var tsObj = ts?.AsObject();
                if (tsObj != null)
                {
                    var position = tsObj["position"]?.GetValue<double>() ?? 0;
                    var tabAlignment = tsObj["alignment"]?.GetValue<string>() ?? "left";
                    var leader = tsObj["leader"]?.GetValue<string>() ?? "none";
                    paraFormat.TabStops.Add(new TabStop(position, GetTabAlignment(tabAlignment), GetTabLeader(leader)));
                }
            }
        }

        if (!string.IsNullOrEmpty(text))
        {
            para.RemoveAllChildren();
            var newRun = new Run(doc, text);
            FontHelper.Word.ApplyFontSettings(newRun, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic,
                underlineStr, color);
            para.AppendChild(newRun);
        }
        else
        {
            var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            if (runs.Count == 0)
            {
                var hasFontSettings = fontName != null || fontNameAscii != null || fontNameFarEast != null ||
                                      fontSize.HasValue || bold.HasValue || italic.HasValue || underlineStr != null ||
                                      color != null;

                if (hasFontSettings)
                {
                    var sentinelRun = new Run(doc, "\u200B");
                    FontHelper.Word.ApplyFontSettings(sentinelRun, fontName, fontNameAscii, fontNameFarEast, fontSize,
                        bold, italic, underlineStr, color);
                    para.AppendChild(sentinelRun);
                }
            }
            else
            {
                foreach (var run in runs)
                    FontHelper.Word.ApplyFontSettings(run, fontName, fontNameAscii, fontNameFarEast, fontSize, bold,
                        italic, underlineStr, color);
            }
        }

        ctx.Save(outputPath);

        var resultMsg = $"Paragraph {paragraphIndex.Value} format edited successfully";
        if (!string.IsNullOrEmpty(text)) resultMsg += ", text content updated";
        return resultMsg;
    }

    /// <summary>
    ///     Gets all paragraphs from the document with optional filtering as JSON.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sectionIndex">The section index to get paragraphs from (0-based).</param>
    /// <param name="includeEmpty">Whether to include empty paragraphs.</param>
    /// <param name="styleFilter">Filter paragraphs by style name.</param>
    /// <param name="includeCommentParagraphs">Whether to include paragraphs inside nested structures.</param>
    /// <param name="includeTextboxParagraphs">Whether to include paragraphs inside TextBox/Shape objects.</param>
    /// <returns>A JSON string containing paragraph information.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex is out of range.</exception>
    private static string GetParagraphs(DocumentContext<Document> ctx, int? sectionIndex, bool includeEmpty,
        string? styleFilter, bool includeCommentParagraphs, bool includeTextboxParagraphs)
    {
        var doc = ctx.Document;

        List<Paragraph> paragraphs;
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            paragraphs = doc.Sections[sectionIndex.Value].Body
                .GetChildNodes(NodeType.Paragraph, includeCommentParagraphs).Cast<Paragraph>().ToList();
        }
        else
        {
            if (includeCommentParagraphs)
            {
                paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            }
            else
            {
                paragraphs = [];
                foreach (var section in doc.Sections.Cast<Section>())
                {
                    var bodyParagraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>()
                        .ToList();
                    paragraphs.AddRange(bodyParagraphs);
                }
            }
        }

        if (!includeEmpty) paragraphs = paragraphs.Where(p => !string.IsNullOrWhiteSpace(p.GetText())).ToList();
        if (!string.IsNullOrEmpty(styleFilter))
            paragraphs = paragraphs.Where(p => p.ParagraphFormat.Style?.Name == styleFilter).ToList();

        if (!includeTextboxParagraphs)
            paragraphs = paragraphs.Where(p =>
            {
                var shapeAncestor = p.GetAncestor(NodeType.Shape);
                if (shapeAncestor is Shape { ShapeType: ShapeType.TextBox }) return false;
                var currentNode = p.ParentNode;
                while (currentNode != null)
                {
                    if (currentNode.NodeType == NodeType.Shape)
                        if (currentNode is Shape { ShapeType: ShapeType.TextBox })
                            return false;
                    currentNode = currentNode.ParentNode;
                }

                return true;
            }).ToList();

        List<object> paragraphList = [];
        for (var i = 0; i < paragraphs.Count; i++)
        {
            var para = paragraphs[i];
            var text = para.GetText().Trim();
            var location = "Body";
            string? commentInfo = null;

            if (para.ParentNode != null)
            {
                var commentAncestor = para.GetAncestor(NodeType.Comment);
                if (commentAncestor != null)
                {
                    location = "Comment";
                    if (commentAncestor is Comment comment)
                        commentInfo = $"ID: {comment.Id}, Author: {comment.Author}";
                }
                else
                {
                    var shapeAncestor = para.GetAncestor(NodeType.Shape);
                    if (shapeAncestor != null)
                    {
                        location = shapeAncestor is Shape { ShapeType: ShapeType.TextBox } ? "TextBox" : "Shape";
                    }
                    else
                    {
                        var bodyAncestor = para.GetAncestor(NodeType.Body);
                        if (bodyAncestor == null || para.ParentNode.NodeType != NodeType.Body)
                            location = para.ParentNode.NodeType.ToString();
                    }
                }
            }

            var paraInfo = new Dictionary<string, object?>
            {
                ["index"] = i,
                ["location"] = location,
                ["style"] = para.ParagraphFormat.Style?.Name,
                ["text"] = text.Length > 100 ? text[..100] + "..." : text,
                ["textLength"] = text.Length
            };

            if (commentInfo != null) paraInfo["commentInfo"] = commentInfo;
            paragraphList.Add(paraInfo);
        }

        var result = new
        {
            count = paragraphs.Count,
            filters = new
                { sectionIndex, includeEmpty, styleFilter, includeCommentParagraphs, includeTextboxParagraphs },
            paragraphs = paragraphList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Gets detailed formatting information for a specific paragraph as JSON.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="paragraphIndex">The paragraph index to get formatting for (0-based).</param>
    /// <param name="includeRunDetails">Whether to include detailed run-level formatting.</param>
    /// <returns>A JSON string containing paragraph formatting information.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraphIndex is null or out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the paragraph cannot be found.</exception>
    private static string GetParagraphFormat(DocumentContext<Document> ctx, int? paragraphIndex, bool includeRunDetails)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for get_format operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        var para = paragraphs[paragraphIndex.Value] as Paragraph;
        if (para == null)
            throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");

        var format = para.ParagraphFormat;
        var text = para.GetText().Trim();

        var resultDict = new Dictionary<string, object?>
        {
            ["paragraphIndex"] = paragraphIndex.Value,
            ["text"] = text,
            ["textLength"] = text.Length,
            ["runCount"] = para.Runs.Count,
            ["paragraphFormat"] = new
            {
                styleName = format.StyleName,
                alignment = format.Alignment.ToString(),
                leftIndent = Math.Round(format.LeftIndent, 2),
                rightIndent = Math.Round(format.RightIndent, 2),
                firstLineIndent = Math.Round(format.FirstLineIndent, 2),
                spaceBefore = Math.Round(format.SpaceBefore, 2),
                spaceAfter = Math.Round(format.SpaceAfter, 2),
                lineSpacing = Math.Round(format.LineSpacing, 2),
                lineSpacingRule = format.LineSpacingRule.ToString()
            }
        };

        if (para.ListFormat is { IsListItem: true })
            resultDict["listFormat"] = new
            {
                isListItem = true,
                listLevel = para.ListFormat.ListLevelNumber,
                listId = para.ListFormat.List?.ListId
            };

        var borders = new Dictionary<string, object>();
        if (format.Borders.Top.LineStyle != LineStyle.None)
            borders["top"] = new
            {
                lineStyle = format.Borders.Top.LineStyle.ToString(), lineWidth = format.Borders.Top.LineWidth,
                color = format.Borders.Top.Color.Name
            };
        if (format.Borders.Bottom.LineStyle != LineStyle.None)
            borders["bottom"] = new
            {
                lineStyle = format.Borders.Bottom.LineStyle.ToString(), lineWidth = format.Borders.Bottom.LineWidth,
                color = format.Borders.Bottom.Color.Name
            };
        if (format.Borders.Left.LineStyle != LineStyle.None)
            borders["left"] = new
            {
                lineStyle = format.Borders.Left.LineStyle.ToString(), lineWidth = format.Borders.Left.LineWidth,
                color = format.Borders.Left.Color.Name
            };
        if (format.Borders.Right.LineStyle != LineStyle.None)
            borders["right"] = new
            {
                lineStyle = format.Borders.Right.LineStyle.ToString(), lineWidth = format.Borders.Right.LineWidth,
                color = format.Borders.Right.Color.Name
            };
        if (borders.Count > 0)
            resultDict["borders"] = borders;

        if (format.Shading.BackgroundPatternColor.ToArgb() != Color.Empty.ToArgb())
        {
            var bgColor = format.Shading.BackgroundPatternColor;
            resultDict["backgroundColor"] = $"#{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
        }

        if (format.TabStops.Count > 0)
        {
            List<object> tabStopsList = [];
            for (var i = 0; i < format.TabStops.Count; i++)
            {
                var tab = format.TabStops[i];
                tabStopsList.Add(new
                {
                    position = Math.Round(tab.Position, 2), alignment = tab.Alignment.ToString(),
                    leader = tab.Leader.ToString()
                });
            }

            resultDict["tabStops"] = tabStopsList;
        }

        if (para.Runs.Count > 0)
        {
            var firstRun = para.Runs[0];
            var fontInfo = new Dictionary<string, object?> { ["fontSize"] = firstRun.Font.Size };

            if (firstRun.Font.NameAscii != firstRun.Font.NameFarEast)
            {
                fontInfo["fontAscii"] = firstRun.Font.NameAscii;
                fontInfo["fontFarEast"] = firstRun.Font.NameFarEast;
            }
            else
            {
                fontInfo["font"] = firstRun.Font.Name;
            }

            if (firstRun.Font.Bold) fontInfo["bold"] = true;
            if (firstRun.Font.Italic) fontInfo["italic"] = true;
            if (firstRun.Font.Underline != Underline.None) fontInfo["underline"] = firstRun.Font.Underline.ToString();
            if (firstRun.Font.StrikeThrough) fontInfo["strikethrough"] = true;
            if (firstRun.Font.Superscript) fontInfo["superscript"] = true;
            if (firstRun.Font.Subscript) fontInfo["subscript"] = true;
            if (firstRun.Font.Color.ToArgb() != Color.Empty.ToArgb())
                fontInfo["color"] = $"#{firstRun.Font.Color.R:X2}{firstRun.Font.Color.G:X2}{firstRun.Font.Color.B:X2}";
            if (firstRun.Font.HighlightColor != Color.Empty)
                fontInfo["highlightColor"] = firstRun.Font.HighlightColor.Name;

            resultDict["fontFormat"] = fontInfo;
        }

        if (includeRunDetails && para.Runs.Count > 1)
        {
            List<object> runs = [];
            for (var i = 0; i < Math.Min(para.Runs.Count, 10); i++)
            {
                var run = para.Runs[i];
                var runInfo = new Dictionary<string, object?>
                {
                    ["index"] = i,
                    ["text"] = run.Text.Replace("\r", "\\r").Replace("\n", "\\n"),
                    ["fontSize"] = run.Font.Size
                };

                if (run.Font.NameAscii != run.Font.NameFarEast)
                {
                    runInfo["fontAscii"] = run.Font.NameAscii;
                    runInfo["fontFarEast"] = run.Font.NameFarEast;
                }
                else
                {
                    runInfo["font"] = run.Font.Name;
                }

                if (run.Font.Bold) runInfo["bold"] = true;
                if (run.Font.Italic) runInfo["italic"] = true;
                if (run.Font.Underline != Underline.None) runInfo["underline"] = run.Font.Underline.ToString();

                runs.Add(runInfo);
            }

            resultDict["runs"] = new
                { total = para.Runs.Count, displayed = Math.Min(para.Runs.Count, 10), details = runs };
        }

        return JsonSerializer.Serialize(resultDict, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Copies formatting from one paragraph to another.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sourceParagraphIndex">The source paragraph index to copy from (0-based).</param>
    /// <param name="targetParagraphIndex">The target paragraph index to copy to (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when paragraphs cannot be retrieved.</exception>
    private static string CopyParagraphFormat(DocumentContext<Document> ctx, string? outputPath,
        int? sourceParagraphIndex, int? targetParagraphIndex)
    {
        if (!sourceParagraphIndex.HasValue)
            throw new ArgumentException("sourceParagraphIndex parameter is required for copy_format operation");
        if (!targetParagraphIndex.HasValue)
            throw new ArgumentException("targetParagraphIndex parameter is required for copy_format operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (sourceParagraphIndex.Value < 0 || sourceParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Source paragraph index {sourceParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (targetParagraphIndex.Value < 0 || targetParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Target paragraph index {targetParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        var sourcePara = paragraphs[sourceParagraphIndex.Value] as Paragraph;
        var targetPara = paragraphs[targetParagraphIndex.Value] as Paragraph;

        if (sourcePara == null || targetPara == null)
            throw new InvalidOperationException("Unable to get paragraphs");

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
        for (var i = 0; i < sourcePara.ParagraphFormat.TabStops.Count; i++)
        {
            var tabStop = sourcePara.ParagraphFormat.TabStops[i];
            targetPara.ParagraphFormat.TabStops.Add(tabStop.Position, tabStop.Alignment, tabStop.Leader);
        }

        ctx.Save(outputPath);

        var result = "Paragraph format copied successfully\n";
        result += $"Source paragraph: #{sourceParagraphIndex.Value}\n";
        result += $"Target paragraph: #{targetParagraphIndex.Value}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Merges multiple consecutive paragraphs into one.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="startParagraphIndex">The starting paragraph index (0-based).</param>
    /// <param name="endParagraphIndex">The ending paragraph index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are invalid.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the start paragraph cannot be found.</exception>
    private static string MergeParagraphs(DocumentContext<Document> ctx, string? outputPath, int? startParagraphIndex,
        int? endParagraphIndex)
    {
        if (!startParagraphIndex.HasValue)
            throw new ArgumentException("startParagraphIndex parameter is required for merge operation");
        if (!endParagraphIndex.HasValue)
            throw new ArgumentException("endParagraphIndex parameter is required for merge operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (startParagraphIndex.Value < 0 || startParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Start paragraph index {startParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (endParagraphIndex.Value < 0 || endParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"End paragraph index {endParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (startParagraphIndex.Value > endParagraphIndex.Value)
            throw new ArgumentException(
                $"Start paragraph index {startParagraphIndex.Value} cannot be greater than end paragraph index {endParagraphIndex.Value}");

        if (startParagraphIndex.Value == endParagraphIndex.Value)
            throw new ArgumentException("Start and end paragraph indices are the same, no merge needed");

        var startPara = paragraphs[startParagraphIndex.Value] as Paragraph;
        if (startPara == null) throw new InvalidOperationException("Unable to get start paragraph");

        for (var i = startParagraphIndex.Value + 1; i <= endParagraphIndex.Value; i++)
            if (paragraphs[i] is Paragraph para)
            {
                if (startPara.Runs.Count > 0)
                {
                    var spaceRun = new Run(doc, " ");
                    startPara.AppendChild(spaceRun);
                }

                var runsToMove = para.Runs.ToArray();
                foreach (var run in runsToMove) startPara.AppendChild(run);

                para.Remove();
            }

        ctx.Save(outputPath);

        var result = "Paragraphs merged successfully\n";
        result += $"Merge range: Paragraph #{startParagraphIndex.Value} to #{endParagraphIndex.Value}\n";
        result += $"Merged paragraphs: {endParagraphIndex.Value - startParagraphIndex.Value + 1}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Converts an alignment string to ParagraphAlignment enum
    /// </summary>
    /// <param name="alignment">Alignment string (left, center, right, justify)</param>
    /// <returns>Corresponding ParagraphAlignment value</returns>
    private static ParagraphAlignment GetAlignment(string alignment)
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

    /// <summary>
    ///     Converts a line spacing rule string to LineSpacingRule enum
    /// </summary>
    /// <param name="rule">Line spacing rule string (atleast, exactly, or default multiple)</param>
    /// <returns>Corresponding LineSpacingRule value</returns>
    private static LineSpacingRule GetLineSpacingRule(string rule)
    {
        return rule.ToLower() switch
        {
            "atleast" => LineSpacingRule.AtLeast,
            "exactly" => LineSpacingRule.Exactly,
            _ => LineSpacingRule.Multiple
        };
    }

    /// <summary>
    ///     Converts a tab alignment string to TabAlignment enum
    /// </summary>
    /// <param name="alignment">Tab alignment string (left, center, right, decimal, bar, clear)</param>
    /// <returns>Corresponding TabAlignment value</returns>
    private static TabAlignment GetTabAlignment(string alignment)
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

    /// <summary>
    ///     Converts a tab leader string to TabLeader enum
    /// </summary>
    /// <param name="leader">Tab leader string (none, dots, dashes, line, heavy, middledot)</param>
    /// <returns>Corresponding TabLeader value</returns>
    private static TabLeader GetTabLeader(string leader)
    {
        return leader.ToLower() switch
        {
            "none" => TabLeader.None,
            "dots" => TabLeader.Dots,
            "dashes" => TabLeader.Dashes,
            "line" => TabLeader.Line,
            "heavy" => TabLeader.Heavy,
            "middledot" => TabLeader.MiddleDot,
            _ => TabLeader.None
        };
    }
}