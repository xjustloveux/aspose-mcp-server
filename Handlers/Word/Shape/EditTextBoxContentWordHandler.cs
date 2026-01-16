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
    /// <exception cref="ArgumentException">Thrown when textboxIndex is missing or out of range.</exception>
    /// <exception cref="Exception">Thrown when textbox paragraph cannot be retrieved.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditTextBoxContentParameters(parameters);

        var doc = context.Document;
        var textboxes = WordShapeHelper.FindAllTextboxes(doc);

        if (p.TextboxIndex < 0 || p.TextboxIndex >= textboxes.Count)
            throw new ArgumentException(
                $"Textbox index {p.TextboxIndex} out of range (total textboxes: {textboxes.Count})");

        var textbox = textboxes[p.TextboxIndex];

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

        if (p.Text != null)
        {
            if (p.AppendText && runsCollection.Count > 0)
            {
                var newRun = new Run(doc, p.Text);
                para.AppendChild(newRun);
            }
            else
            {
                para.RemoveAllChildren();
                var newRun = new Run(doc, p.Text);
                para.AppendChild(newRun);
            }

            runs = para.GetChildNodes(NodeType.Run, false).Cast<Run>().ToList();
        }

        if (p.ClearFormatting)
            foreach (var run in runs)
                run.Font.ClearFormatting();

        var hasFormatting = !string.IsNullOrEmpty(p.FontName) || !string.IsNullOrEmpty(p.FontNameAscii) ||
                            !string.IsNullOrEmpty(p.FontNameFarEast) || p.FontSize.HasValue ||
                            p.Bold.HasValue || p.Italic.HasValue || !string.IsNullOrEmpty(p.Color);

        if (hasFormatting)
            foreach (var run in runs)
            {
                FontHelper.Word.ApplyFontSettings(
                    run,
                    p.FontName,
                    p.FontNameAscii,
                    p.FontNameFarEast,
                    p.FontSize,
                    p.Bold,
                    p.Italic
                );

                if (!string.IsNullOrEmpty(p.Color))
                    run.Font.Color = ColorHelper.ParseColor(p.Color, true);
            }

        MarkModified(context);

        return $"Successfully edited textbox #{p.TextboxIndex}.";
    }

    private static EditTextBoxContentParameters ExtractEditTextBoxContentParameters(OperationParameters parameters)
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

        return new EditTextBoxContentParameters(
            textboxIndex.Value,
            text,
            appendText,
            fontName,
            fontNameAscii,
            fontNameFarEast,
            fontSize,
            bold,
            italic,
            color,
            clearFormatting);
    }

    private record EditTextBoxContentParameters(
        int TextboxIndex,
        string? Text,
        bool AppendText,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        string? Color,
        bool ClearFormatting);
}
