using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for adding textboxes to Word documents.
/// </summary>
public class AddTextBoxWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_textbox";

    /// <summary>
    ///     Adds a textbox to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: textboxWidth, textboxHeight, positionX, positionY, backgroundColor, borderColor,
    ///     borderWidth, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, textAlignment
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var textBoxParams = ExtractTextBoxParameters(parameters);
        if (string.IsNullOrEmpty(textBoxParams.Text))
            throw new ArgumentException("text is required for add_textbox operation");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var textBox = CreateTextBoxShape(doc, textBoxParams);
        ApplyTextBoxStyles(textBox, textBoxParams);

        var para = CreateTextParagraph(doc, textBoxParams);
        textBox.AppendChild(para);
        builder.InsertNode(textBox);

        MarkModified(context);
        return "Successfully added textbox.";
    }

    /// <summary>
    ///     Extracts textbox parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted textbox parameters.</returns>
    private static TextBoxParameters ExtractTextBoxParameters(OperationParameters parameters)
    {
        return new TextBoxParameters(
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional("textboxWidth", 200.0),
            parameters.GetOptional("textboxHeight", 100.0),
            parameters.GetOptional("positionX", 100.0),
            parameters.GetOptional("positionY", 100.0),
            parameters.GetOptional<string?>("backgroundColor"),
            parameters.GetOptional<string?>("borderColor"),
            parameters.GetOptional("borderWidth", 1.0),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<string?>("fontNameAscii"),
            parameters.GetOptional<string?>("fontNameFarEast"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional("textAlignment", "left")
        );
    }

    /// <summary>
    ///     Creates the textbox shape.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="p">The textbox parameters.</param>
    /// <returns>The created textbox shape.</returns>
    private static Aspose.Words.Drawing.Shape CreateTextBoxShape(Document doc, TextBoxParameters p)
    {
        return new Aspose.Words.Drawing.Shape(doc, ShapeType.TextBox)
        {
            Width = p.TextboxWidth,
            Height = p.TextboxHeight,
            Left = p.PositionX,
            Top = p.PositionY,
            WrapType = WrapType.None,
            RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
            RelativeVerticalPosition = RelativeVerticalPosition.Page
        };
    }

    /// <summary>
    ///     Applies styles to the textbox shape.
    /// </summary>
    /// <param name="textBox">The textbox shape.</param>
    /// <param name="p">The textbox parameters.</param>
    private static void ApplyTextBoxStyles(Aspose.Words.Drawing.Shape textBox, TextBoxParameters p)
    {
        if (!string.IsNullOrEmpty(p.BackgroundColor))
        {
            textBox.Fill.Color = ColorHelper.ParseColor(p.BackgroundColor);
            textBox.Fill.Visible = true;
        }

        if (!string.IsNullOrEmpty(p.BorderColor))
        {
            textBox.Stroke.Color = ColorHelper.ParseColor(p.BorderColor);
            textBox.Stroke.Weight = p.BorderWidth;
            textBox.Stroke.Visible = true;
        }
    }

    /// <summary>
    ///     Creates the text paragraph for the textbox.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="p">The textbox parameters.</param>
    /// <returns>The created text paragraph.</returns>
    private static WordParagraph CreateTextParagraph(Document doc, TextBoxParameters p)
    {
        var para = new WordParagraph(doc);
        var run = new Run(doc, p.Text);

        ApplyFontSettings(run, p);
        para.ParagraphFormat.Alignment = WordShapeHelper.ParseAlignment(p.TextAlignment);
        para.AppendChild(run);

        return para;
    }

    /// <summary>
    ///     Applies font settings to a run.
    /// </summary>
    /// <param name="run">The run to apply settings to.</param>
    /// <param name="p">The textbox parameters.</param>
    private static void ApplyFontSettings(Run run, TextBoxParameters p)
    {
        if (!string.IsNullOrEmpty(p.FontNameAscii))
            run.Font.NameAscii = p.FontNameAscii;

        if (!string.IsNullOrEmpty(p.FontNameFarEast))
            run.Font.NameFarEast = p.FontNameFarEast;

        if (!string.IsNullOrEmpty(p.FontName))
            ApplyFontName(run, p);

        if (p.FontSize.HasValue)
            run.Font.Size = p.FontSize.Value;

        if (p.Bold.HasValue)
            run.Font.Bold = p.Bold.Value;
    }

    /// <summary>
    ///     Applies font name settings to a run.
    /// </summary>
    /// <param name="run">The run to apply settings to.</param>
    /// <param name="p">The textbox parameters.</param>
    private static void ApplyFontName(Run run, TextBoxParameters p)
    {
        if (string.IsNullOrEmpty(p.FontNameAscii) && string.IsNullOrEmpty(p.FontNameFarEast))
        {
            run.Font.Name = p.FontName;
        }
        else
        {
            if (string.IsNullOrEmpty(p.FontNameAscii))
                run.Font.NameAscii = p.FontName;
            if (string.IsNullOrEmpty(p.FontNameFarEast))
                run.Font.NameFarEast = p.FontName;
        }
    }

    /// <summary>
    ///     Record to hold textbox creation parameters.
    /// </summary>
    private sealed record TextBoxParameters(
        string? Text,
        double TextboxWidth,
        double TextboxHeight,
        double PositionX,
        double PositionY,
        string? BackgroundColor,
        string? BorderColor,
        double BorderWidth,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        string TextAlignment);
}
