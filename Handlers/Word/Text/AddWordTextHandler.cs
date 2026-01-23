using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for adding text to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds text to the end of the document with optional formatting.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text.
    ///     Optional: fontName, fontSize, bold, italic, underline, color, strikethrough, superscript, subscript.
    /// </param>
    /// <returns>Success message with formatting details.</returns>
    /// <exception cref="ArgumentException">Thrown when text parameter is missing.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddWordTextParameters(parameters);

        var doc = context.Document;
        doc.EnsureMinimum();
        var lastSection = doc.LastSection;
        var body = lastSection.Body;

        var lines = p.Text.Contains('\n') || p.Text.Contains('\r')
            ? p.Text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None)
            : [p.Text];

        var builder = new DocumentBuilder(doc);
        MoveToBodyEnd(builder, body);

        foreach (var line in lines)
        {
            var currentParaBefore = builder.CurrentParagraph;
            var needsNewParagraph = false;
            if (currentParaBefore != null)
            {
                var existingRuns = currentParaBefore.GetChildNodes(NodeType.Run, false);
                var existingText = currentParaBefore.GetText().Trim();
                needsNewParagraph = existingRuns.Count > 0 || !string.IsNullOrEmpty(existingText);
            }

            if (needsNewParagraph)
            {
                builder.Writeln();
                builder.MoveTo(builder.CurrentParagraph);
            }

            ClearFormatting(builder);
            ApplyFontFormatting(builder, p);

            builder.Write(line);

            ApplyRunFormatting(builder.CurrentParagraph, line, p);
        }

        MarkModified(context);

        return BuildResultMessage(p);
    }

    /// <summary>
    ///     Extracts add word text parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add word text parameters.</returns>
    private static AddWordTextParameters ExtractAddWordTextParameters(OperationParameters parameters)
    {
        return new AddWordTextParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<string?>("underline"),
            parameters.GetOptional<string?>("color"),
            parameters.GetOptional<bool?>("strikethrough"),
            parameters.GetOptional<bool?>("superscript"),
            parameters.GetOptional<bool?>("subscript"));
    }

    /// <summary>
    ///     Moves the document builder to the end of the body, avoiding shapes/textboxes.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="body">The document body.</param>
    private static void MoveToBodyEnd(DocumentBuilder builder, Body body)
    {
        MoveToLastParagraph(builder, body);
        EscapeFromShapeIfNeeded(builder, body);
    }

    /// <summary>
    ///     Moves to the last paragraph in the body.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="body">The document body.</param>
    private static void MoveToLastParagraph(DocumentBuilder builder, Body body)
    {
        var bodyParagraphs = body.GetChildNodes(NodeType.Paragraph, false);
        if (bodyParagraphs.Count > 0 && bodyParagraphs[^1] is WordParagraph lastBodyPara)
            builder.MoveTo(lastBodyPara);
        else
            builder.MoveToDocumentEnd();
    }

    /// <summary>
    ///     Escapes from a shape/textbox if the builder is inside one.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="body">The document body.</param>
    private static void EscapeFromShapeIfNeeded(DocumentBuilder builder, Body body)
    {
        var currentNode = builder.CurrentNode;
        if (currentNode == null) return;

        var shapeAncestor = currentNode.GetAncestor(NodeType.Shape);
        if (shapeAncestor == null) return;

        var bodyParagraphs = body.GetChildNodes(NodeType.Paragraph, false);
        if (bodyParagraphs.Count > 0 && bodyParagraphs[^1] is WordParagraph lastBodyPara)
            builder.MoveTo(lastBodyPara);
        else
            builder.MoveTo(body);
    }

    /// <summary>
    ///     Clears all formatting from the document builder.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    private static void ClearFormatting(DocumentBuilder builder)
    {
        builder.Font.ClearFormatting();
        builder.Font.Bold = false;
        builder.Font.Italic = false;
        builder.Font.Underline = Underline.None;
        builder.Font.StrikeThrough = false;
        builder.Font.Superscript = false;
        builder.Font.Subscript = false;
        builder.ParagraphFormat.ClearFormatting();
    }

    /// <summary>
    ///     Applies font formatting to the document builder.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="p">The text parameters containing font settings.</param>
    private static void ApplyFontFormatting(DocumentBuilder builder, AddWordTextParameters p)
    {
        FontHelper.Word.ApplyFontSettings(
            builder,
            p.FontName,
            fontSize: p.FontSize,
            bold: p.Bold,
            italic: p.Italic,
            underline: p.Underline,
            color: p.Color,
            strikethrough: p.Strikethrough,
            superscript: p.Superscript,
            subscript: p.Subscript
        );
    }

    /// <summary>
    ///     Applies formatting to runs created by the builder.
    /// </summary>
    /// <param name="para">The paragraph containing the runs.</param>
    /// <param name="line">The text line to match.</param>
    /// <param name="p">The text parameters containing font settings.</param>
    private static void ApplyRunFormatting(WordParagraph? para, string line, AddWordTextParameters p)
    {
        if (para == null) return;

        var runs = para.GetChildNodes(NodeType.Run, false);
        foreach (var node in runs)
            if (node is Run run && run.Text == line)
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
                    bold: p.Bold,
                    italic: p.Italic,
                    underline: p.Underline,
                    strikethrough: p.Strikethrough,
                    superscript: p.Superscript,
                    subscript: p.Subscript
                );
            }
    }

    /// <summary>
    ///     Builds the result message with applied formatting details.
    /// </summary>
    /// <param name="p">The text parameters containing formatting settings.</param>
    /// <returns>The formatted result message.</returns>
    private static SuccessResult BuildResultMessage(AddWordTextParameters p)
    {
        List<string> formatInfo = [];
        if (p.Bold == true) formatInfo.Add("bold");
        if (p.Italic == true) formatInfo.Add("italic");
        if (!string.IsNullOrEmpty(p.Underline) && p.Underline != "none") formatInfo.Add($"underline({p.Underline})");
        if (p.Strikethrough == true) formatInfo.Add("strikethrough");
        if (p.Superscript == true) formatInfo.Add("superscript");
        if (p.Subscript == true) formatInfo.Add("subscript");

        var message = "Text added to document successfully.";
        if (formatInfo.Count > 0)
            message += $" Applied formats: {string.Join(", ", formatInfo)}.";

        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Record to hold add word text parameters.
    /// </summary>
    /// <param name="Text">The text to add.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Bold">Whether to apply bold.</param>
    /// <param name="Italic">Whether to apply italic.</param>
    /// <param name="Underline">The underline style.</param>
    /// <param name="Color">The font color.</param>
    /// <param name="Strikethrough">Whether to apply strikethrough.</param>
    /// <param name="Superscript">Whether to apply superscript.</param>
    /// <param name="Subscript">Whether to apply subscript.</param>
    private sealed record AddWordTextParameters(
        string Text,
        string? FontName,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        string? Underline,
        string? Color,
        bool? Strikethrough,
        bool? Superscript,
        bool? Subscript);
}
