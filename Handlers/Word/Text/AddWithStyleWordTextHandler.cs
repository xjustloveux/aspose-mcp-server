using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for adding text with style to Word documents.
/// </summary>
public class AddWithStyleWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_with_style";

    /// <summary>
    ///     Adds text with specified style and formatting.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text.
    ///     Optional: styleName, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic,
    ///     underline, color, alignment, indentLevel, leftIndent, firstLineIndent, tabStops, paragraphIndexForAdd.
    /// </param>
    /// <returns>Success message with details.</returns>
    /// <exception cref="ArgumentException">Thrown when text is missing or style is not found.</exception>
    /// <exception cref="InvalidOperationException">Thrown when style cannot be applied.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddWithStyleParameters(parameters);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        JsonArray? tabStops = null;
        if (!string.IsNullOrEmpty(p.TabStopsJson))
            tabStops = JsonNode.Parse(p.TabStopsJson) as JsonArray;

        var (targetPara, warningMessage) = SetupBuilderPosition(doc, builder, p.ParagraphIndex);

        var para = CreateStyledParagraph(doc, p.Text, p.StyleName, p.FontName, p.FontNameAscii, p.FontNameFarEast,
            p.FontSize, p.Bold, p.Italic, p.Underline, p.Color, p.Alignment, p.IndentLevel, p.LeftIndent,
            p.FirstLineIndent, tabStops);

        InsertParagraph(builder, para, targetPara, p.ParagraphIndex);
        FixEmptyParagraphStyles(doc, para);

        MarkModified(context);

        return BuildResultMessage(p.ParagraphIndex, p.StyleName, p.FontName, p.FontNameAscii, p.FontNameFarEast,
            p.FontSize, p.Bold, p.Italic, p.Underline, p.Color, p.Alignment, p.IndentLevel, p.LeftIndent,
            p.FirstLineIndent, warningMessage);
    }

    private static AddWithStyleParameters ExtractAddWithStyleParameters(OperationParameters parameters)
    {
        return new AddWithStyleParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<string?>("styleName"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<bool?>("underline"),
            parameters.GetOptional<string?>("color"),
            parameters.GetOptional<string?>("alignment"),
            parameters.GetOptional<int?>("indentLevel"),
            parameters.GetOptional<double?>("leftIndent"),
            parameters.GetOptional<double?>("firstLineIndent"),
            parameters.GetOptional<string?>("tabStops"),
            parameters.GetOptional<int?>("paragraphIndexForAdd"));
    }

    /// <summary>
    ///     Sets up the document builder position based on paragraph index.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="builder">The document builder.</param>
    /// <param name="paragraphIndex">The optional paragraph index.</param>
    /// <returns>Tuple of target paragraph and warning message.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when paragraph cannot be found.</exception>
    private static (WordParagraph? targetPara, string warningMessage) SetupBuilderPosition(Document doc,
        DocumentBuilder builder, int? paragraphIndex)
    {
        WordParagraph? targetPara = null;
        var warningMessage = "";

        if (!paragraphIndex.HasValue)
        {
            builder.MoveToDocumentEnd();
            return (targetPara, warningMessage);
        }

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphIndex.Value == -1)
        {
            if (paragraphs.Count > 0 && paragraphs[0] is WordParagraph firstPara)
            {
                targetPara = firstPara;
                builder.MoveTo(targetPara);
            }
        }
        else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
        {
            if (paragraphs[paragraphIndex.Value] is WordParagraph para)
            {
                targetPara = para;
                builder.MoveTo(targetPara);
            }
            else
            {
                throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");
            }
        }
        else
        {
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
        }

        return (targetPara, warningMessage);
    }

    /// <summary>
    ///     Creates a styled paragraph with the specified formatting.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="text">The text content.</param>
    /// <param name="styleName">The style name.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size.</param>
    /// <param name="bold">Whether text should be bold.</param>
    /// <param name="italic">Whether text should be italic.</param>
    /// <param name="underline">Whether text should be underlined.</param>
    /// <param name="color">The text color.</param>
    /// <param name="alignment">The paragraph alignment.</param>
    /// <param name="indentLevel">The indentation level.</param>
    /// <param name="leftIndent">The left indentation in points.</param>
    /// <param name="firstLineIndent">The first line indentation in points.</param>
    /// <param name="tabStops">Custom tab stops.</param>
    /// <returns>The created paragraph.</returns>
    private static WordParagraph CreateStyledParagraph(Document doc, string text, string? styleName,
        string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize,
        bool? bold, bool? italic, bool? underline, string? color, string? alignment,
        int? indentLevel, double? leftIndent, double? firstLineIndent, JsonArray? tabStops)
    {
        var para = new WordParagraph(doc);
        var run = new Run(doc, text);

        ApplyStyle(doc, para, styleName);
        ApplyFontFormatting(run, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline, color);
        ApplyParagraphFormatting(para, alignment, indentLevel, leftIndent, firstLineIndent);
        ApplyTabStops(para, tabStops);

        para.AppendChild(run);
        return para;
    }

    /// <summary>
    ///     Applies style to the paragraph.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The paragraph.</param>
    /// <param name="styleName">The style name.</param>
    /// <exception cref="InvalidOperationException">Thrown when style is not found or cannot be applied.</exception>
    private static void ApplyStyle(Document doc, WordParagraph para, string? styleName)
    {
        if (string.IsNullOrEmpty(styleName)) return;

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
    }

    /// <summary>
    ///     Applies font formatting to the run.
    /// </summary>
    private static void ApplyFontFormatting(Run run, string? fontName, string? fontNameAscii,
        string? fontNameFarEast, double? fontSize, bool? bold, bool? italic, bool? underline, string? color)
    {
        var underlineStr = underline.HasValue ? underline.Value ? "single" : "none" : null;
        FontHelper.Word.ApplyFontSettings(run, fontName, fontNameAscii, fontNameFarEast, fontSize,
            bold, italic, underlineStr, color);
    }

    /// <summary>
    ///     Applies paragraph formatting.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="alignment">The alignment.</param>
    /// <param name="indentLevel">The indent level.</param>
    /// <param name="leftIndent">The left indentation.</param>
    /// <param name="firstLineIndent">The first line indentation.</param>
    private static void ApplyParagraphFormatting(WordParagraph para, string? alignment, int? indentLevel,
        double? leftIndent, double? firstLineIndent)
    {
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
        else if (leftIndent.HasValue)
            para.ParagraphFormat.LeftIndent = leftIndent.Value;

        if (firstLineIndent.HasValue)
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
    }

    /// <summary>
    ///     Applies custom tab stops to the paragraph.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="tabStops">The tab stops JSON array.</param>
    private static void ApplyTabStops(WordParagraph para, JsonArray? tabStops)
    {
        if (tabStops is not { Count: > 0 }) return;

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

    /// <summary>
    ///     Inserts the paragraph at the appropriate position.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="para">The paragraph to insert.</param>
    /// <param name="targetPara">The target paragraph for positioning.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    private static void InsertParagraph(DocumentBuilder builder, WordParagraph? para, WordParagraph? targetPara,
        int? paragraphIndex)
    {
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
    }

    /// <summary>
    ///     Fixes empty paragraphs created after insertion to use Normal style.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The inserted paragraph.</param>
    private static void FixEmptyParagraphStyles(Document doc, WordParagraph para)
    {
        var parentNode = para.ParentNode;
        if (parentNode == null) return;

        var allParagraphs = parentNode.GetChildNodes(NodeType.Paragraph, false).Cast<WordParagraph>().ToList();
        var insertedIndex = allParagraphs.IndexOf(para);

        for (var i = insertedIndex + 1; i < allParagraphs.Count; i++)
        {
            var nextPara = allParagraphs[i];
            if (!string.IsNullOrWhiteSpace(nextPara.GetText())) break;

            try
            {
                var normalStyle = doc.Styles[StyleIdentifier.Normal];
                if (normalStyle != null)
                {
                    nextPara.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                    nextPara.ParagraphFormat.Style = normalStyle;
                    nextPara.ParagraphFormat.StyleName = "Normal";
                    nextPara.ParagraphFormat.ClearFormatting();
                    nextPara.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                    nextPara.ParagraphFormat.Style = normalStyle;
                    nextPara.ParagraphFormat.StyleName = "Normal";
                }
            }
            catch
            {
                try
                {
                    nextPara.ParagraphFormat.ClearFormatting();
                    nextPara.ParagraphFormat.StyleName = "Normal";
                    nextPara.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
                }
                catch
                {
                    // Skip if unable to set style
                }
            }
        }
    }

    /// <summary>
    ///     Builds the result message with details of the operation.
    /// </summary>
    private static string BuildResultMessage(int? paragraphIndex, string? styleName, string? fontName,
        string? fontNameAscii, string? fontNameFarEast, double? fontSize, bool? bold, bool? italic,
        bool? underline, string? color, string? alignment, int? indentLevel, double? leftIndent,
        double? firstLineIndent, string warningMessage)
    {
        var result = "Text added successfully.";

        if (paragraphIndex.HasValue)
            result += paragraphIndex.Value == -1
                ? " Insert position: beginning of document."
                : $" Insert position: after paragraph #{paragraphIndex.Value}.";
        else
            result += " Insert position: end of document.";

        if (!string.IsNullOrEmpty(styleName))
        {
            result += $" Applied style: {styleName}.";
        }
        else
        {
            var customFormatting = new List<string>();
            if (!string.IsNullOrEmpty(fontNameAscii)) customFormatting.Add($"Font (ASCII): {fontNameAscii}");
            if (!string.IsNullOrEmpty(fontNameFarEast)) customFormatting.Add($"Font (Far East): {fontNameFarEast}");
            if (!string.IsNullOrEmpty(fontName) && string.IsNullOrEmpty(fontNameAscii) &&
                string.IsNullOrEmpty(fontNameFarEast))
                customFormatting.Add($"Font: {fontName}");
            if (fontSize.HasValue) customFormatting.Add($"Font size: {fontSize.Value} pt");
            if (bold == true) customFormatting.Add("Bold");
            if (italic == true) customFormatting.Add("Italic");
            if (underline == true) customFormatting.Add("Underline");
            if (!string.IsNullOrEmpty(color)) customFormatting.Add($"Color: {color}");
            if (!string.IsNullOrEmpty(alignment)) customFormatting.Add($"Alignment: {alignment}");

            if (customFormatting.Count > 0)
                result += $" Custom formatting: {string.Join(", ", customFormatting)}.";
        }

        if (indentLevel.HasValue)
            result += $" Indent level: {indentLevel.Value} ({indentLevel.Value * 36} pt).";
        else if (leftIndent.HasValue)
            result += $" Left indent: {leftIndent.Value} pt.";

        if (firstLineIndent.HasValue)
            result += $" First line indent: {firstLineIndent.Value} pt.";

        if (!string.IsNullOrEmpty(warningMessage))
            result += warningMessage;

        return Success(result);
    }

    private record AddWithStyleParameters(
        string Text,
        string? StyleName,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        string? Color,
        string? Alignment,
        int? IndentLevel,
        double? LeftIndent,
        double? FirstLineIndent,
        string? TabStopsJson,
        int? ParagraphIndex);
}
