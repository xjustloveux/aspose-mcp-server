using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for editing textbox content in Word documents.
/// </summary>
public class EditTextBoxContentWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit_textbox_content";

    /// <summary>
    ///     Edits the content of a textbox.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: textboxIndex
    ///     Optional: text, appendText, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, color,
    ///     clearFormatting
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var textboxIndex = parameters.GetOptional<int?>("textboxIndex");
        var text = parameters.GetOptional<string?>("text");
        var appendText = parameters.GetOptional("appendText", false);
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontNameAscii = parameters.GetOptional<string?>("fontNameAscii");
        var fontNameFarEast = parameters.GetOptional<string?>("fontNameFarEast");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var color = parameters.GetOptional<string?>("color");
        var clearFormatting = parameters.GetOptional("clearFormatting", false);

        if (!textboxIndex.HasValue)
            throw new ArgumentException("textboxIndex is required for edit_textbox_content operation");

        var doc = context.Document;
        var textboxes = WordShapeHelper.FindAllTextboxes(doc);

        if (textboxIndex.Value < 0 || textboxIndex.Value >= textboxes.Count)
            throw new ArgumentException(
                $"Textbox index {textboxIndex.Value} out of range (total textboxes: {textboxes.Count})");

        var textbox = textboxes[textboxIndex.Value];

        var paragraphs = textbox.GetChildNodes(NodeType.Paragraph, false);
        WordParagraph para;

        if (paragraphs.Count == 0)
        {
            para = new WordParagraph(doc);
            textbox.AppendChild(para);
        }
        else
        {
            para = paragraphs[0] as WordParagraph ?? throw new Exception("Cannot get textbox paragraph");
        }

        var runsCollection = para.GetChildNodes(NodeType.Run, false);
        var runs = runsCollection.Cast<Run>().ToList();

        if (text != null)
        {
            if (appendText && runsCollection.Count > 0)
            {
                var newRun = new Run(doc, text);
                para.AppendChild(newRun);
            }
            else
            {
                para.RemoveAllChildren();
                var newRun = new Run(doc, text);
                para.AppendChild(newRun);
            }

            runs = para.GetChildNodes(NodeType.Run, false).Cast<Run>().ToList();
        }

        if (clearFormatting)
            foreach (var run in runs)
                run.Font.ClearFormatting();

        var hasFormatting = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) ||
                            !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue ||
                            bold.HasValue || italic.HasValue || !string.IsNullOrEmpty(color);

        if (hasFormatting)
            foreach (var run in runs)
            {
                FontHelper.Word.ApplyFontSettings(
                    run,
                    fontName,
                    fontNameAscii,
                    fontNameFarEast,
                    fontSize,
                    bold,
                    italic
                );

                // Handle color separately to throw exception on parse error
                if (!string.IsNullOrEmpty(color))
                    run.Font.Color = ColorHelper.ParseColor(color, true);
            }

        MarkModified(context);

        return $"Successfully edited textbox #{textboxIndex.Value}.";
    }
}
