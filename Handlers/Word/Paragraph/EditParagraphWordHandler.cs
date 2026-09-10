using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for editing paragraphs in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EditParagraphWordHandler : OperationHandlerBase<Document>
{
    private const string SingleLineSpacingRule = "single";

    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits paragraph content and formatting.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    ///     Optional: sectionIndex, text, styleName, alignment, fontName, fontNameAscii, fontNameFarEast, fontSize,
    ///     bold, italic, underline, color, indentLeft, indentRight, firstLineIndent, spaceBefore, spaceAfter,
    ///     lineSpacing, lineSpacingRule, tabStops
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var editParams = ExtractEditParagraphParameters(parameters);

        if (!editParams.ParagraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for edit operation");

        var doc = context.Document;

        var paragraphRef =
            ParagraphResolver.Resolve(doc, ParagraphAddress.From(parameters, editParams.ParagraphIndex.Value));
        var para = paragraphRef.Paragraph;
        var resolvedIndex = paragraphRef.Address.Index;

        // Everything that can refuse this request is decided before anything is changed. The
        // applies below run in sequence — font, formatting, line spacing, style, tab stops, text —
        // so a refusal from a later one used to leave the earlier ones applied: measured, an edit
        // refused for a field left the paragraph's alignment changed from Left to Center, and
        // `IsModified` was still false because MarkModified runs last (R10-W01).
        ValidateBeforeMutating(doc, para, editParams);

        var builder = new DocumentBuilder(doc);
        builder.MoveTo(para.FirstChild ?? para);

        var underlineStr = GetUnderlineString(editParams.Underline);
        var fontSettings = new FontParams(editParams.FontName, editParams.FontNameAscii, editParams.FontNameFarEast,
            editParams.FontSize, editParams.Bold, editParams.Italic,
            underlineStr, editParams.Color);

        FontHelper.Word.ApplyFontSettings(builder, editParams.FontName, editParams.FontNameAscii,
            editParams.FontNameFarEast, editParams.FontSize, editParams.Bold, editParams.Italic,
            underlineStr, editParams.Color);

        ApplyParagraphFormatting(para, editParams.Alignment, editParams.IndentLeft, editParams.IndentRight,
            editParams.FirstLineIndent, editParams.SpaceBefore, editParams.SpaceAfter);
        ApplyLineSpacing(para.ParagraphFormat, editParams.LineSpacing, editParams.LineSpacingRule);
        ApplyStyle(doc, para, editParams.StyleName);
        ApplyTabStops(para.ParagraphFormat, editParams.TabStops);
        ApplyTextContent(doc, para, editParams.Text, fontSettings);

        MarkModified(context);

        var resultMsg = $"Paragraph {resolvedIndex} format edited successfully";
        if (!string.IsNullOrEmpty(editParams.Text)) resultMsg += ", text content updated";
        return new SuccessResult { Message = resultMsg };
    }

    /// <summary>
    ///     Resolves every value this edit depends on, and refuses the whole request if any of them
    ///     cannot be resolved — before a single property is written.
    /// </summary>
    /// <param name="doc">The document being edited.</param>
    /// <param name="para">The paragraph being edited.</param>
    /// <param name="editParams">The requested edit.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when a style, alignment, spacing rule, tab stop or the text replacement itself
    ///     cannot be applied. Every one of these used to be discovered partway through, leaving
    ///     the paragraph half-edited by an operation that reported failure (R10-W01).
    /// </exception>
    private static void ValidateBeforeMutating(Document doc, Aspose.Words.Paragraph para,
        EditParagraphParameters editParams)
    {
        if (!string.IsNullOrEmpty(editParams.Alignment))
            WordParagraphHelper.GetAlignment(editParams.Alignment);

        if (editParams.LineSpacing.HasValue || !string.IsNullOrEmpty(editParams.LineSpacingRule))
            WordParagraphHelper.GetLineSpacingRule(editParams.LineSpacingRule ?? SingleLineSpacingRule);

        if (!string.IsNullOrEmpty(editParams.StyleName) && doc.Styles[editParams.StyleName] == null)
            throw new ArgumentException(
                $"Style '{editParams.StyleName}' not found. Use word_get_styles tool to view "
                + "available styles");

        // The whole array, every field. Checking only alignment and leader left `position` to be
        // read inside ApplyTabStops — which clears the paragraph's existing stops first, so a
        // malformed position threw after the old stops were already gone (§23.4).
        WordTabStopHelper.Resolve(editParams.TabStops);

        if (string.IsNullOrEmpty(editParams.Text)) return;

        // The same question ReplaceParagraphText asks, asked before anything is written.
        var extents = FieldBoundaryHelper.FieldExtents.Of(doc);
        var textRuns = WordRunHelper.GetDirectRuns(para)
            .Where(run => extents.EnclosingField(run) == null).ToList();

        if (textRuns.Count == 0
            && (extents.EnclosingField(para) != null || extents.WouldSplitAField(para)))
            throw new ArgumentException(
                "This paragraph lies inside a field, so replacing its text would write into the "
                + "field's code or result rather than into the document. Edit the field with "
                + "word_field, or choose a paragraph outside it.");
    }

    /// <summary>
    ///     Applies paragraph formatting options.
    /// </summary>
    /// <param name="para">The paragraph to format.</param>
    /// <param name="alignment">The text alignment.</param>
    /// <param name="indentLeft">The left indent.</param>
    /// <param name="indentRight">The right indent.</param>
    /// <param name="firstLineIndent">The first line indent.</param>
    /// <param name="spaceBefore">The space before.</param>
    /// <param name="spaceAfter">The space after.</param>
    private static void ApplyParagraphFormatting(Aspose.Words.Paragraph para, string? alignment, double? indentLeft,
        double? indentRight, double? firstLineIndent, double? spaceBefore, double? spaceAfter)
    {
        var paraFormat = para.ParagraphFormat;

        if (!string.IsNullOrEmpty(alignment))
            paraFormat.Alignment = WordParagraphHelper.GetAlignment(alignment);

        if (indentLeft.HasValue) paraFormat.LeftIndent = indentLeft.Value;
        if (indentRight.HasValue) paraFormat.RightIndent = indentRight.Value;
        if (firstLineIndent.HasValue) paraFormat.FirstLineIndent = firstLineIndent.Value;
        if (spaceBefore.HasValue) paraFormat.SpaceBefore = spaceBefore.Value;
        if (spaceAfter.HasValue) paraFormat.SpaceAfter = spaceAfter.Value;
    }

    /// <summary>
    ///     Applies line spacing settings.
    /// </summary>
    /// <param name="paraFormat">The paragraph format.</param>
    /// <param name="lineSpacing">The line spacing value.</param>
    /// <param name="lineSpacingRule">The line spacing rule.</param>
    private static void ApplyLineSpacing(ParagraphFormat paraFormat, double? lineSpacing, string? lineSpacingRule)
    {
        if (!lineSpacing.HasValue && string.IsNullOrEmpty(lineSpacingRule)) return;

        var effectiveRule = lineSpacingRule ?? SingleLineSpacingRule;
        var rule = WordParagraphHelper.GetLineSpacingRule(effectiveRule);
        paraFormat.LineSpacingRule = rule;

        paraFormat.LineSpacing = lineSpacing ?? GetDefaultLineSpacing(effectiveRule);
    }

    /// <summary>
    ///     Gets the default line spacing value for a rule.
    /// </summary>
    /// <param name="rule">The line spacing rule.</param>
    /// <returns>The default line spacing value.</returns>
    private static double GetDefaultLineSpacing(string rule)
    {
        return rule.ToLower() switch
        {
            SingleLineSpacingRule => 1.0,
            "oneandhalf" => 1.5,
            "double" => 2.0,
            _ => 1.0
        };
    }

    /// <summary>
    ///     Applies a style to the paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="para">The paragraph.</param>
    /// <param name="styleName">The style name.</param>
    /// <exception cref="ArgumentException">Thrown when style is not found.</exception>
    private static void ApplyStyle(Document doc, Aspose.Words.Paragraph para, string? styleName)
    {
        if (string.IsNullOrEmpty(styleName)) return;

        var style = doc.Styles[styleName];
        if (style == null)
            throw new ArgumentException(
                $"Style '{styleName}' not found. Use word_get_styles tool to view available styles");

        var isEmpty = string.IsNullOrWhiteSpace(para.GetText());
        if (isEmpty) para.ParagraphFormat.ClearFormatting();
        para.ParagraphFormat.Style = style;
        para.ParagraphFormat.StyleName = styleName;
    }

    /// <summary>
    ///     Applies tab stops to the paragraph format.
    /// </summary>
    /// <param name="paraFormat">The paragraph format.</param>
    /// <param name="tabStops">The tab stops array.</param>
    private static void ApplyTabStops(ParagraphFormat paraFormat, JsonArray? tabStops)
    {
        if (tabStops is not { Count: > 0 }) return;

        // Resolved in full before the existing stops are cleared. Reading the values inside this
        // loop meant a malformed one threw after Clear(), leaving the paragraph with no stops at
        // all and the request reporting failure (§23.4). The reading itself lives in
        // WordTabStopHelper, because three other handlers were reading the same array by their own
        // rules (R13-W01).
        var resolved = WordTabStopHelper.Resolve(tabStops);

        paraFormat.TabStops.Clear();
        foreach (var stop in resolved) paraFormat.TabStops.Add(stop);
    }

    /// <summary>
    ///     Applies text content to the paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="para">The paragraph.</param>
    /// <param name="text">The text content.</param>
    /// <param name="fontParams">The font parameters.</param>
    private static void ApplyTextContent(Document doc, Aspose.Words.Paragraph para, string? text, FontParams fontParams)
    {
        if (!string.IsNullOrEmpty(text))
        {
            ReplaceParagraphText(doc, para, text, fontParams);
            return;
        }

        var runs = WordRunHelper.GetDirectRuns(para);
        if (runs.Count == 0)
        {
            if (fontParams.HasAnySettings())
            {
                var sentinelRun = new Run(doc, "\u200B");
                FontHelper.Word.ApplyFontSettings(sentinelRun, fontParams.FontName, fontParams.FontNameAscii,
                    fontParams.FontNameFarEast, fontParams.FontSize, fontParams.Bold, fontParams.Italic,
                    fontParams.UnderlineStr, fontParams.Color);
                para.AppendChild(sentinelRun);
            }
        }
        else
        {
            foreach (var run in runs)
                FontHelper.Word.ApplyFontSettings(run, fontParams.FontName, fontParams.FontNameAscii,
                    fontParams.FontNameFarEast, fontParams.FontSize, fontParams.Bold, fontParams.Italic,
                    fontParams.UnderlineStr, fontParams.Color);
        }
    }

    /// <summary>
    ///     Replaces the paragraph's plain-text runs with a single new run carrying the supplied
    ///     text, while preserving any fields, bookmarks, and inline objects in place. The new run
    ///     takes the position of the first replaced text run; when the paragraph contains only
    ///     field content it is appended after it.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The paragraph whose text is replaced.</param>
    /// <param name="text">The new text content.</param>
    /// <param name="fontParams">The font parameters to apply to the new run.</param>
    private static void ReplaceParagraphText(Document doc, Aspose.Words.Paragraph para, string text,
        FontParams fontParams)
    {
        var newRun = new Run(doc, text);
        FontHelper.Word.ApplyFontSettings(newRun, fontParams.FontName, fontParams.FontNameAscii,
            fontParams.FontNameFarEast, fontParams.FontSize, fontParams.Bold, fontParams.Italic,
            fontParams.UnderlineStr, fontParams.Color);

        // Indexed once: asking per run would number the whole document per run.
        var extents = FieldBoundaryHelper.FieldExtents.Of(para.Document as Document
                                                          ?? throw new InvalidOperationException(
                                                              "The paragraph is not part of a document."));
        var textRuns = WordRunHelper.GetDirectRuns(para)
            .Where(run => extents.EnclosingField(run) == null).ToList();

        if (textRuns.Count == 0)
        {
            // Every run in this paragraph is inside a field, so there is nowhere in it that is
            // ordinary text. Appending anyway put the caller's text inside the field's result or
            // its code — for a field spanning several paragraphs, a whole paragraph can be in
            // that position (§21.3).
            if (extents.EnclosingField(para) != null || extents.WouldSplitAField(para))
                throw new ArgumentException(
                    "This paragraph lies inside a field, so replacing its text would write into "
                    + "the field's code or result rather than into the document. Edit the field "
                    + "with word_field, or choose a paragraph outside it.");

            para.AppendChild(newRun);
        }
        else
        {
            textRuns[0].ParentNode.InsertBefore(newRun, textRuns[0]);
        }

        foreach (var run in textRuns)
            run.Remove();
    }

    /// <summary>
    ///     Extracts edit paragraph parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit paragraph parameters.</returns>
    private static EditParagraphParameters ExtractEditParagraphParameters(OperationParameters parameters)
    {
        return new EditParagraphParameters(
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional<string?>("styleName"),
            parameters.GetOptional<string?>("alignment"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<bool?>("underline"),
            parameters.GetOptional<string?>("color"),
            parameters.GetOptional<double?>("indentLeft"),
            parameters.GetOptional<double?>("indentRight"),
            parameters.GetOptional<double?>("firstLineIndent"),
            parameters.GetOptional<double?>("spaceBefore"),
            parameters.GetOptional<double?>("spaceAfter"),
            parameters.GetOptional<double?>("lineSpacing"),
            parameters.GetOptional<string?>("lineSpacingRule"),
            parameters.GetOptional<JsonArray?>("tabStops")
        );
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
    ///     Record to hold edit paragraph parameters.
    /// </summary>
    /// <param name="ParagraphIndex">The paragraph index to edit (-1 for last).</param>
    /// <param name="Text">The text content.</param>
    /// <param name="StyleName">The style name.</param>
    /// <param name="Alignment">The text alignment.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Bold">Whether text is bold.</param>
    /// <param name="Italic">Whether text is italic.</param>
    /// <param name="Underline">Whether text is underlined.</param>
    /// <param name="Color">The text color.</param>
    /// <param name="IndentLeft">The left indent.</param>
    /// <param name="IndentRight">The right indent.</param>
    /// <param name="FirstLineIndent">The first line indent.</param>
    /// <param name="SpaceBefore">The space before.</param>
    /// <param name="SpaceAfter">The space after.</param>
    /// <param name="LineSpacing">The line spacing value.</param>
    /// <param name="LineSpacingRule">The line spacing rule.</param>
    /// <param name="TabStops">The tab stops array.</param>
    private sealed record EditParagraphParameters(
        int? ParagraphIndex,
        string? Text,
        string? StyleName,
        string? Alignment,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        bool? Underline,
        string? Color,
        double? IndentLeft,
        double? IndentRight,
        double? FirstLineIndent,
        double? SpaceBefore,
        double? SpaceAfter,
        double? LineSpacing,
        string? LineSpacingRule,
        JsonArray? TabStops);

    /// <summary>
    ///     Record to hold font parameters.
    /// </summary>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Bold">Whether text is bold.</param>
    /// <param name="Italic">Whether text is italic.</param>
    /// <param name="UnderlineStr">The underline style string.</param>
    /// <param name="Color">The text color.</param>
    private sealed record FontParams(
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        string? UnderlineStr,
        string? Color)
    {
        /// <summary>
        ///     Checks if any font settings are specified.
        /// </summary>
        /// <returns>True if any settings are specified.</returns>
        public bool HasAnySettings()
        {
            return FontName != null || FontNameAscii != null || FontNameFarEast != null ||
                   FontSize.HasValue || Bold.HasValue || Italic.HasValue || UnderlineStr != null || Color != null;
        }
    }
}
