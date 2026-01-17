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

        var para = CreateStyledParagraph(doc, p, tabStops);

        InsertParagraph(builder, para, targetPara, p.ParagraphIndex);
        FixEmptyParagraphStyles(doc, para);

        MarkModified(context);

        return BuildResultMessage(p, warningMessage);
    }

    /// <summary>
    ///     Extracts add with style parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add with style parameters.</returns>
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
    /// <param name="p">The parameters containing text and formatting settings.</param>
    /// <param name="tabStops">Custom tab stops.</param>
    /// <returns>The created paragraph.</returns>
    private static WordParagraph CreateStyledParagraph(Document doc, AddWithStyleParameters p, JsonArray? tabStops)
    {
        var para = new WordParagraph(doc);
        var run = new Run(doc, p.Text);

        ApplyStyle(doc, para, p.StyleName);
        ApplyFontFormatting(run, p);
        ApplyParagraphFormatting(para, p);
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
    /// <exception cref="ArgumentException">Thrown when the style is not found.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the style cannot be applied.</exception>
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
    /// <param name="run">The run to format.</param>
    /// <param name="p">The parameters containing font settings.</param>
    private static void ApplyFontFormatting(Run run, AddWithStyleParameters p)
    {
        var underlineStr = GetUnderlineString(p.Underline);
        FontHelper.Word.ApplyFontSettings(run, p.FontName, p.FontNameAscii, p.FontNameFarEast, p.FontSize,
            p.Bold, p.Italic, underlineStr, p.Color);
    }

    /// <summary>
    ///     Converts nullable bool underline value to string representation.
    /// </summary>
    /// <param name="underline">The nullable underline value.</param>
    /// <returns>The underline string: "single", "none", or null.</returns>
    private static string? GetUnderlineString(bool? underline)
    {
        if (!underline.HasValue) return null;
        return underline.Value ? "single" : "none";
    }

    /// <summary>
    ///     Applies paragraph formatting.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="p">The parameters containing paragraph formatting settings.</param>
    private static void ApplyParagraphFormatting(WordParagraph para, AddWithStyleParameters p)
    {
        if (!string.IsNullOrEmpty(p.Alignment))
            para.ParagraphFormat.Alignment = p.Alignment.ToLower() switch
            {
                "left" => ParagraphAlignment.Left,
                "right" => ParagraphAlignment.Right,
                "center" => ParagraphAlignment.Center,
                "justify" => ParagraphAlignment.Justify,
                _ => ParagraphAlignment.Left
            };

        if (p.IndentLevel.HasValue)
            para.ParagraphFormat.LeftIndent = p.IndentLevel.Value * 36;
        else if (p.LeftIndent.HasValue)
            para.ParagraphFormat.LeftIndent = p.LeftIndent.Value;

        if (p.FirstLineIndent.HasValue)
            para.ParagraphFormat.FirstLineIndent = p.FirstLineIndent.Value;
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
    /// <param name="p">The parameters containing formatting settings.</param>
    /// <param name="warningMessage">Any warning message to append.</param>
    /// <returns>The formatted result message.</returns>
    private static string
        BuildResultMessage(AddWithStyleParameters p,
            string warningMessage) // NOSONAR S3776 - Linear parameter validation sequence
    {
        var result = "Text added successfully.";

        if (p.ParagraphIndex.HasValue)
            result += p.ParagraphIndex.Value == -1
                ? " Insert position: beginning of document."
                : $" Insert position: after paragraph #{p.ParagraphIndex.Value}.";
        else
            result += " Insert position: end of document.";

        if (!string.IsNullOrEmpty(p.StyleName))
        {
            result += $" Applied style: {p.StyleName}.";
        }
        else
        {
            var customFormatting = new List<string>();
            if (!string.IsNullOrEmpty(p.FontNameAscii)) customFormatting.Add($"Font (ASCII): {p.FontNameAscii}");
            if (!string.IsNullOrEmpty(p.FontNameFarEast)) customFormatting.Add($"Font (Far East): {p.FontNameFarEast}");
            if (!string.IsNullOrEmpty(p.FontName) && string.IsNullOrEmpty(p.FontNameAscii) &&
                string.IsNullOrEmpty(p.FontNameFarEast))
                customFormatting.Add($"Font: {p.FontName}");
            if (p.FontSize.HasValue) customFormatting.Add($"Font size: {p.FontSize.Value} pt");
            if (p.Bold == true) customFormatting.Add("Bold");
            if (p.Italic == true) customFormatting.Add("Italic");
            if (p.Underline == true) customFormatting.Add("Underline");
            if (!string.IsNullOrEmpty(p.Color)) customFormatting.Add($"Color: {p.Color}");
            if (!string.IsNullOrEmpty(p.Alignment)) customFormatting.Add($"Alignment: {p.Alignment}");

            if (customFormatting.Count > 0)
                result += $" Custom formatting: {string.Join(", ", customFormatting)}.";
        }

        if (p.IndentLevel.HasValue)
            result += $" Indent level: {p.IndentLevel.Value} ({p.IndentLevel.Value * 36} pt).";
        else if (p.LeftIndent.HasValue)
            result += $" Left indent: {p.LeftIndent.Value} pt.";

        if (p.FirstLineIndent.HasValue)
            result += $" First line indent: {p.FirstLineIndent.Value} pt.";

        if (!string.IsNullOrEmpty(warningMessage))
            result += warningMessage;

        return Success(result);
    }

    /// <summary>
    ///     Record to hold add with style parameters.
    /// </summary>
    /// <param name="Text">The text to add.</param>
    /// <param name="StyleName">The style name.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Bold">Whether to apply bold.</param>
    /// <param name="Italic">Whether to apply italic.</param>
    /// <param name="Underline">Whether to apply underline.</param>
    /// <param name="Color">The font color.</param>
    /// <param name="Alignment">The paragraph alignment.</param>
    /// <param name="IndentLevel">The indent level.</param>
    /// <param name="LeftIndent">The left indent in points.</param>
    /// <param name="FirstLineIndent">The first line indent in points.</param>
    /// <param name="TabStopsJson">The tab stops JSON string.</param>
    /// <param name="ParagraphIndex">The paragraph index for insertion.</param>
    private sealed record AddWithStyleParameters(
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
