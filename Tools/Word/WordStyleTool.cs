using System.ComponentModel;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing styles in Word documents (get, create, apply, copy)
/// </summary>
[McpServerToolType]
public class WordStyleTool
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
    ///     Initializes a new instance of the WordStyleTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordStyleTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word style operation (get_styles, create_style, apply_style, copy_styles).
    /// </summary>
    /// <param name="operation">The operation to perform: get_styles, create_style, apply_style, copy_styles.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="includeBuiltIn">Include built-in styles (for get_styles, default: false).</param>
    /// <param name="styleName">Style name (for create_style, apply_style).</param>
    /// <param name="styleType">Style type: paragraph, character, table, list (for create_style, default: paragraph).</param>
    /// <param name="baseStyle">Base style to inherit from (for create_style).</param>
    /// <param name="fontName">Font name (for create_style).</param>
    /// <param name="fontNameAscii">Font name for ASCII characters (for create_style).</param>
    /// <param name="fontNameFarEast">Font name for Far East characters (for create_style).</param>
    /// <param name="fontSize">Font size in points (for create_style).</param>
    /// <param name="bold">Bold text (for create_style).</param>
    /// <param name="italic">Italic text (for create_style).</param>
    /// <param name="underline">Underline text (for create_style).</param>
    /// <param name="color">Text color hex (for create_style).</param>
    /// <param name="alignment">Paragraph alignment: left, center, right, justify (for create_style).</param>
    /// <param name="spaceBefore">Space before paragraph in points (for create_style).</param>
    /// <param name="spaceAfter">Space after paragraph in points (for create_style).</param>
    /// <param name="lineSpacing">Line spacing multiplier (for create_style).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, for apply_style).</param>
    /// <param name="paragraphIndices">Array of paragraph indices (for apply_style).</param>
    /// <param name="sectionIndex">Section index (0-based, for apply_style, default: 0).</param>
    /// <param name="tableIndex">Table index (0-based, for apply_style).</param>
    /// <param name="applyToAllParagraphs">Apply to all paragraphs (for apply_style, default: false).</param>
    /// <param name="sourceDocument">Source document path to copy styles from (for copy_styles).</param>
    /// <param name="styleNames">Array of style names to copy (for copy_styles).</param>
    /// <param name="overwriteExisting">Overwrite existing styles (for copy_styles, default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get_styles.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_style")]
    [Description(
        @"Manage styles in Word documents. Supports 4 operations: get_styles, create_style, apply_style, copy_styles.

Usage examples:
- Get styles: word_style(operation='get_styles', path='doc.docx', includeBuiltIn=true)
- Create style: word_style(operation='create_style', path='doc.docx', styleName='CustomStyle', styleType='paragraph', fontSize=14, bold=true)
- Apply style: word_style(operation='apply_style', path='doc.docx', styleName='Heading 1', paragraphIndex=0)
- Copy styles: word_style(operation='copy_styles', path='doc.docx', sourceDocument='template.docx')")]
    public string Execute(
        [Description("Operation: get_styles, create_style, apply_style, copy_styles")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Include built-in styles (for get_styles, default: false)")]
        bool includeBuiltIn = false,
        [Description("Style name (for create_style, apply_style)")]
        string? styleName = null,
        [Description("Style type: paragraph, character, table, list (for create_style, default: paragraph)")]
        string styleType = "paragraph",
        [Description("Base style to inherit from (for create_style)")]
        string? baseStyle = null,
        [Description("Font name (for create_style)")]
        string? fontName = null,
        [Description("Font name for ASCII characters (for create_style)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (for create_style)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (for create_style)")]
        double? fontSize = null,
        [Description("Bold text (for create_style)")]
        bool? bold = null,
        [Description("Italic text (for create_style)")]
        bool? italic = null,
        [Description("Underline text (for create_style)")]
        bool? underline = null,
        [Description("Text color hex (for create_style)")]
        string? color = null,
        [Description("Paragraph alignment: left, center, right, justify (for create_style)")]
        string? alignment = null,
        [Description("Space before paragraph in points (for create_style)")]
        double? spaceBefore = null,
        [Description("Space after paragraph in points (for create_style)")]
        double? spaceAfter = null,
        [Description("Line spacing multiplier (for create_style)")]
        double? lineSpacing = null,
        [Description("Paragraph index (0-based, for apply_style)")]
        int? paragraphIndex = null,
        [Description("Array of paragraph indices (for apply_style)")]
        int[]? paragraphIndices = null,
        [Description("Section index (0-based, for apply_style, default: 0)")]
        int sectionIndex = 0,
        [Description("Table index (0-based, for apply_style)")]
        int? tableIndex = null,
        [Description("Apply to all paragraphs (for apply_style, default: false)")]
        bool applyToAllParagraphs = false,
        [Description("Source document path to copy styles from (for copy_styles)")]
        string? sourceDocument = null,
        [Description("Array of style names to copy (for copy_styles)")]
        string[]? styleNames = null,
        [Description("Overwrite existing styles (for copy_styles, default: false)")]
        bool overwriteExisting = false)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "get_styles" => GetStyles(ctx, includeBuiltIn),
            "create_style" => CreateStyle(ctx, outputPath, styleName, styleType, baseStyle, fontName, fontNameAscii,
                fontNameFarEast, fontSize, bold, italic, underline, color, alignment, spaceBefore, spaceAfter,
                lineSpacing),
            "apply_style" => ApplyStyle(ctx, outputPath, styleName, paragraphIndex, paragraphIndices, sectionIndex,
                tableIndex, applyToAllParagraphs),
            "copy_styles" => CopyStyles(ctx, outputPath, sourceDocument, styleNames, overwriteExisting),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets all styles from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="includeBuiltIn">Whether to include built-in styles.</param>
    /// <returns>A JSON string containing style information.</returns>
    private static string GetStyles(DocumentContext<Document> ctx, bool includeBuiltIn)
    {
        var doc = ctx.Document;

        List<Style> paraStyles;

        if (includeBuiltIn)
        {
            // Include all styles (built-in and custom)
            paraStyles = doc.Styles
                .Where(s => s.Type == StyleType.Paragraph)
                .OrderBy(s => s.Name)
                .ToList();
        }
        else
        {
            var usedStyleNames = new HashSet<string>();
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>();
            foreach (var para in paragraphs)
                if (para.ParagraphFormat.Style != null && !string.IsNullOrEmpty(para.ParagraphFormat.Style.Name))
                    usedStyleNames.Add(para.ParagraphFormat.Style.Name);

            // Return styles that are either custom (not built-in) OR are built-in but actually used in the document
            paraStyles = doc.Styles
                .Where(s => s.Type == StyleType.Paragraph && (!s.BuiltIn || usedStyleNames.Contains(s.Name)))
                .OrderBy(s => s.Name)
                .ToList();
        }

        List<object> styleList = [];
        foreach (var style in paraStyles)
        {
            var font = style.Font;
            var paraFormat = style.ParagraphFormat;

            var styleInfo = new Dictionary<string, object?>
            {
                ["name"] = style.Name,
                ["builtIn"] = style.BuiltIn
            };

            if (!string.IsNullOrEmpty(style.BaseStyleName))
                styleInfo["basedOn"] = style.BaseStyleName;

            if (font.NameAscii != font.NameFarEast)
            {
                styleInfo["fontAscii"] = font.NameAscii;
                styleInfo["fontFarEast"] = font.NameFarEast;
            }
            else
            {
                styleInfo["font"] = font.Name;
            }

            styleInfo["fontSize"] = font.Size;
            if (font.Bold) styleInfo["bold"] = true;
            if (font.Italic) styleInfo["italic"] = true;

            styleInfo["alignment"] = paraFormat.Alignment.ToString();
            if (paraFormat.SpaceBefore != 0) styleInfo["spaceBefore"] = paraFormat.SpaceBefore;
            if (paraFormat.SpaceAfter != 0) styleInfo["spaceAfter"] = paraFormat.SpaceAfter;

            styleList.Add(styleInfo);
        }

        var result = new
        {
            count = paraStyles.Count,
            includeBuiltIn,
            note = includeBuiltIn
                ? null
                : "Showing custom styles and built-in styles actually used in the document",
            paragraphStyles = styleList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Creates a new style.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="styleName">The name of the style to create.</param>
    /// <param name="styleTypeStr">The style type (paragraph, character, table, list).</param>
    /// <param name="baseStyle">The base style to inherit from.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text should be bold.</param>
    /// <param name="italic">Whether the text should be italic.</param>
    /// <param name="underline">Whether the text should be underlined.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="alignment">The paragraph alignment.</param>
    /// <param name="spaceBefore">The space before paragraph in points.</param>
    /// <param name="spaceAfter">The space after paragraph in points.</param>
    /// <param name="lineSpacing">The line spacing multiplier.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when styleName is null or empty.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the style already exists.</exception>
    private static string CreateStyle(DocumentContext<Document> ctx, string? outputPath, string? styleName,
        string styleTypeStr, string? baseStyle, string? fontName, string? fontNameAscii, string? fontNameFarEast,
        double? fontSize, bool? bold, bool? italic, bool? underline, string? color, string? alignment,
        double? spaceBefore, double? spaceAfter, double? lineSpacing)
    {
        if (string.IsNullOrEmpty(styleName))
            throw new ArgumentException("styleName is required for create_style operation");

        var doc = ctx.Document;

        if (doc.Styles[styleName] != null)
            throw new InvalidOperationException($"Style '{styleName}' already exists");

        var styleType = styleTypeStr.ToLower() switch
        {
            "character" => StyleType.Character,
            "table" => StyleType.Table,
            "list" => StyleType.List,
            _ => StyleType.Paragraph
        };

        var style = doc.Styles.Add(styleType, styleName);

        if (!string.IsNullOrEmpty(baseStyle))
        {
            var baseStyleObj = doc.Styles[baseStyle];
            if (baseStyleObj != null)
                style.BaseStyleName = baseStyle;
            else
                Console.Error.WriteLine(
                    $"[WARN] Base style '{baseStyle}' not found, style will not inherit from it");
        }

        // Font settings are only applicable to Paragraph, Character, and Table styles
        // List styles don't support direct font settings
        if (styleType != StyleType.List)
        {
            if (!string.IsNullOrEmpty(fontNameAscii))
                style.Font.NameAscii = fontNameAscii;

            if (!string.IsNullOrEmpty(fontNameFarEast))
                style.Font.NameFarEast = fontNameFarEast;

            if (!string.IsNullOrEmpty(fontName))
            {
                if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
                {
                    style.Font.Name = fontName;
                }
                else
                {
                    if (string.IsNullOrEmpty(fontNameAscii))
                        style.Font.NameAscii = fontName;
                    if (string.IsNullOrEmpty(fontNameFarEast))
                        style.Font.NameFarEast = fontName;
                }
            }

            if (fontSize.HasValue)
                style.Font.Size = fontSize.Value;

            if (bold.HasValue)
                style.Font.Bold = bold.Value;

            if (italic.HasValue)
                style.Font.Italic = italic.Value;

            if (underline.HasValue)
                style.Font.Underline = underline.Value ? Underline.Single : Underline.None;

            if (!string.IsNullOrEmpty(color))
                style.Font.Color = ColorHelper.ParseColor(color, true);
        }

        if (styleType == StyleType.Paragraph || styleType == StyleType.List)
        {
            if (!string.IsNullOrEmpty(alignment))
                style.ParagraphFormat.Alignment = alignment.ToLower() switch
                {
                    "center" => ParagraphAlignment.Center,
                    "right" => ParagraphAlignment.Right,
                    "justify" => ParagraphAlignment.Justify,
                    _ => ParagraphAlignment.Left
                };

            if (spaceBefore.HasValue)
                style.ParagraphFormat.SpaceBefore = spaceBefore.Value;

            if (spaceAfter.HasValue)
                style.ParagraphFormat.SpaceAfter = spaceAfter.Value;

            if (lineSpacing.HasValue)
            {
                style.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                style.ParagraphFormat.LineSpacing = lineSpacing.Value * 12;
            }
        }

        ctx.Save(outputPath);
        var result = $"Style '{styleName}' created successfully\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Applies a style to paragraphs or runs.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="styleName">The name of the style to apply.</param>
    /// <param name="paragraphIndex">The paragraph index to apply the style to (0-based).</param>
    /// <param name="paragraphIndices">An array of paragraph indices to apply the style to.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="tableIndex">The table index to apply the style to (0-based).</param>
    /// <param name="applyToAllParagraphs">Whether to apply the style to all paragraphs.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when styleName is missing, style is not found, or indices are invalid.</exception>
    private static string ApplyStyle(DocumentContext<Document> ctx, string? outputPath, string? styleName,
        int? paragraphIndex, int[]? paragraphIndices, int sectionIndex, int? tableIndex, bool applyToAllParagraphs)
    {
        if (string.IsNullOrEmpty(styleName))
            throw new ArgumentException("styleName is required for apply_style operation");

        var doc = ctx.Document;
        var style = doc.Styles[styleName];
        if (style == null)
            throw new ArgumentException($"Style '{styleName}' not found");

        var appliedCount = 0;

        if (tableIndex.HasValue)
        {
            var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex.Value < 0 || tableIndex.Value >= tables.Count)
                throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");
            tables[tableIndex.Value].Style = style;
            appliedCount = 1;
        }
        else if (applyToAllParagraphs)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            foreach (var para in paragraphs)
            {
                ApplyStyleToParagraph(para, style, styleName);
                appliedCount++;
            }
        }
        else if (paragraphIndices is { Length: > 0 })
        {
            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[sectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            foreach (var idx in paragraphIndices)
                if (idx >= 0 && idx < paragraphs.Count)
                {
                    ApplyStyleToParagraph(paragraphs[idx], style, styleName);
                    appliedCount++;
                }
        }
        else if (paragraphIndex.HasValue)
        {
            if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[sectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException(
                    $"paragraphIndex must be between 0 and {paragraphs.Count - 1} (section {sectionIndex} has {paragraphs.Count} paragraphs, total document paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count})");

            ApplyStyleToParagraph(paragraphs[paragraphIndex.Value], style, styleName);
            appliedCount = 1;
        }
        else
        {
            throw new ArgumentException(
                "Either paragraphIndex, paragraphIndices, tableIndex, or applyToAllParagraphs must be provided");
        }

        ctx.Save(outputPath);
        var result = $"Applied style '{styleName}' to {appliedCount} element(s)\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Copies styles from source document to destination document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sourceDocument">The source document path to copy styles from.</param>
    /// <param name="styleNames">An array of style names to copy (or null to copy all).</param>
    /// <param name="overwriteExisting">Whether to overwrite existing styles.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sourceDocument is null or empty.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the source document is not found.</exception>
    private static string CopyStyles(DocumentContext<Document> ctx, string? outputPath, string? sourceDocument,
        string[]? styleNames, bool overwriteExisting)
    {
        if (string.IsNullOrEmpty(sourceDocument))
            throw new ArgumentException("sourceDocument is required for copy_styles operation");

        SecurityHelper.ValidateFilePath(sourceDocument, "sourceDocument", true);

        if (!File.Exists(sourceDocument))
            throw new FileNotFoundException($"Source document not found: {sourceDocument}");

        var targetDoc = ctx.Document;
        var sourceDoc = new Document(sourceDocument);

        var styleNamesList = styleNames?.ToList() ?? [];
        var copyAll = styleNamesList.Count == 0;
        var copiedCount = 0;
        var skippedCount = 0;

        foreach (var sourceStyle in sourceDoc.Styles)
        {
            if (!copyAll && !styleNamesList.Contains(sourceStyle.Name))
                continue;

            var existingStyle = targetDoc.Styles[sourceStyle.Name];

            if (existingStyle != null && !overwriteExisting)
            {
                skippedCount++;
                continue;
            }

            try
            {
                if (existingStyle != null && overwriteExisting)
                {
                    CopyStyleProperties(sourceStyle, existingStyle);
                }
                else
                {
                    var newStyle = targetDoc.Styles.Add(sourceStyle.Type, sourceStyle.Name);
                    CopyStyleProperties(sourceStyle, newStyle);
                }

                copiedCount++;
            }
            catch (Exception ex)
            {
                skippedCount++;
                Console.Error.WriteLine($"[WARN] Failed to copy style '{sourceStyle.Name}': {ex.Message}");
            }
        }

        ctx.Save(outputPath);
        var result =
            $"Copied {copiedCount} style(s) from {Path.GetFileName(sourceDocument)}. Skipped: {skippedCount}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Applies a style to a single paragraph, handling empty paragraphs specially.
    /// </summary>
    /// <param name="para">The paragraph to apply the style to.</param>
    /// <param name="style">The style to apply.</param>
    /// <param name="styleName">The name of the style.</param>
    private static void ApplyStyleToParagraph(Paragraph para, Style style, string styleName)
    {
        var paraFormat = para.ParagraphFormat;
        var isEmpty = string.IsNullOrWhiteSpace(para.GetText());

        if (isEmpty)
        {
            paraFormat.ClearFormatting();
            try
            {
                if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                    paraFormat.StyleIdentifier = style.StyleIdentifier;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Failed to set StyleIdentifier for style '{styleName}': {ex.Message}");
            }
        }

        paraFormat.Style = style;
        paraFormat.StyleName = styleName;

        if (isEmpty)
        {
            paraFormat.ClearFormatting();
            paraFormat.Style = style;
            paraFormat.StyleName = styleName;
            try
            {
                if (style.StyleIdentifier != StyleIdentifier.Normal || styleName == "Normal")
                    paraFormat.StyleIdentifier = style.StyleIdentifier;
            }
            catch
            {
                // Ignore StyleIdentifier errors for empty paragraphs
            }
        }
    }

    /// <summary>
    ///     Copies style properties from source to target style.
    /// </summary>
    /// <param name="sourceStyle">The source style to copy from.</param>
    /// <param name="targetStyle">The target style to copy to.</param>
    private static void CopyStyleProperties(Style sourceStyle, Style targetStyle)
    {
        targetStyle.Font.Name = sourceStyle.Font.Name;
        targetStyle.Font.NameAscii = sourceStyle.Font.NameAscii;
        targetStyle.Font.NameFarEast = sourceStyle.Font.NameFarEast;
        targetStyle.Font.Size = sourceStyle.Font.Size;
        targetStyle.Font.Bold = sourceStyle.Font.Bold;
        targetStyle.Font.Italic = sourceStyle.Font.Italic;
        targetStyle.Font.Color = sourceStyle.Font.Color;
        targetStyle.Font.Underline = sourceStyle.Font.Underline;

        if (sourceStyle.Type == StyleType.Paragraph)
        {
            targetStyle.ParagraphFormat.Alignment = sourceStyle.ParagraphFormat.Alignment;
            targetStyle.ParagraphFormat.SpaceBefore = sourceStyle.ParagraphFormat.SpaceBefore;
            targetStyle.ParagraphFormat.SpaceAfter = sourceStyle.ParagraphFormat.SpaceAfter;
            targetStyle.ParagraphFormat.LineSpacing = sourceStyle.ParagraphFormat.LineSpacing;
            targetStyle.ParagraphFormat.LineSpacingRule = sourceStyle.ParagraphFormat.LineSpacingRule;
            targetStyle.ParagraphFormat.LeftIndent = sourceStyle.ParagraphFormat.LeftIndent;
            targetStyle.ParagraphFormat.RightIndent = sourceStyle.ParagraphFormat.RightIndent;
            targetStyle.ParagraphFormat.FirstLineIndent = sourceStyle.ParagraphFormat.FirstLineIndent;
        }
        else if (sourceStyle.Type == StyleType.Table)
        {
            // Copy table-specific properties when available
            try
            {
                targetStyle.ParagraphFormat.Alignment = sourceStyle.ParagraphFormat.Alignment;
            }
            catch
            {
                // Table styles may not support all paragraph properties
            }
        }
    }
}