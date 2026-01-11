using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for editing paragraphs in Word documents.
/// </summary>
public class EditParagraphWordHandler : OperationHandlerBase<Document>
{
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");
        var text = parameters.GetOptional<string?>("text");
        var styleName = parameters.GetOptional<string?>("styleName");
        var alignment = parameters.GetOptional<string?>("alignment");
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontNameAscii = parameters.GetOptional<string?>("fontNameAscii");
        var fontNameFarEast = parameters.GetOptional<string?>("fontNameFarEast");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var underline = parameters.GetOptional<bool?>("underline");
        var color = parameters.GetOptional<string?>("color");
        var indentLeft = parameters.GetOptional<double?>("indentLeft");
        var indentRight = parameters.GetOptional<double?>("indentRight");
        var firstLineIndent = parameters.GetOptional<double?>("firstLineIndent");
        var spaceBefore = parameters.GetOptional<double?>("spaceBefore");
        var spaceAfter = parameters.GetOptional<double?>("spaceAfter");
        var lineSpacing = parameters.GetOptional<double?>("lineSpacing");
        var lineSpacingRule = parameters.GetOptional<string?>("lineSpacingRule");
        var tabStops = parameters.GetOptional<JsonArray?>("tabStops");

        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for edit operation");

        var doc = context.Document;
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
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Aspose.Words.Paragraph>().ToList();

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
            paraFormat.Alignment = WordParagraphHelper.GetAlignment(alignment);

        if (indentLeft.HasValue) paraFormat.LeftIndent = indentLeft.Value;
        if (indentRight.HasValue) paraFormat.RightIndent = indentRight.Value;
        if (firstLineIndent.HasValue) paraFormat.FirstLineIndent = firstLineIndent.Value;
        if (spaceBefore.HasValue) paraFormat.SpaceBefore = spaceBefore.Value;
        if (spaceAfter.HasValue) paraFormat.SpaceAfter = spaceAfter.Value;

        if (lineSpacing.HasValue || !string.IsNullOrEmpty(lineSpacingRule))
        {
            var rule = WordParagraphHelper.GetLineSpacingRule(lineSpacingRule ?? "single");
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
                    paraFormat.TabStops.Add(new TabStop(position,
                        WordParagraphHelper.GetTabAlignment(tabAlignment),
                        WordParagraphHelper.GetTabLeader(leader)));
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

        MarkModified(context);

        var resultMsg = $"Paragraph {paragraphIndex.Value} format edited successfully";
        if (!string.IsNullOrEmpty(text)) resultMsg += ", text content updated";
        return resultMsg;
    }
}
