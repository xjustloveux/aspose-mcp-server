using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Text;

/// <summary>
///     Handler for adding text to PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddPdfTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds text to a specific page in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text.
    ///     Optional: pageIndex (default: 1), x, y, fontName, fontSize, color.
    /// </param>
    /// <returns>Success message with text addition details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        if (string.IsNullOrEmpty(p.Text))
            throw new ArgumentException("text is required");

        var document = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];

        var textFragment = new TextFragment(p.Text) { Position = new Position(p.X, p.Y) };

        var textState = textFragment.TextState;
        FontHelper.Pdf.ApplyFontSettings(textState, p.FontName, p.FontSize);
        textState.ForegroundColor = Color.FromRgb(ColorHelper.ParseColor(p.Color));

        var textBuilder = new TextBuilder(page);
        textBuilder.AppendText(textFragment);

        MarkModified(context);

        return new SuccessResult { Message = $"Text added to page {p.PageIndex}." };
    }

    /// <summary>
    ///     Extracts parameters for add operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional("x", 100.0),
            parameters.GetOptional("y", 700.0),
            parameters.GetOptional("fontName", "Arial"),
            parameters.GetOptional("fontSize", 12.0),
            parameters.GetOptional("color", "Black")
        );
    }

    /// <summary>
    ///     Parameters for add operation.
    /// </summary>
    /// <param name="Text">The text to add.</param>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="X">The X coordinate.</param>
    /// <param name="Y">The Y coordinate.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Color">The text color.</param>
    private sealed record AddParameters(
        string Text,
        int PageIndex,
        double X,
        double Y,
        string FontName,
        double FontSize,
        string Color);
}
