using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Pdf.Text;

/// <summary>
///     Handler for adding text to PDF documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetRequired<string>("text");

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required");

        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var x = parameters.GetOptional("x", 100.0);
        var y = parameters.GetOptional("y", 700.0);
        var fontName = parameters.GetOptional("fontName", "Arial");
        var fontSize = parameters.GetOptional("fontSize", 12.0);
        var color = parameters.GetOptional("color", "Black");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        var textFragment = new TextFragment(text) { Position = new Position(x, y) };

        var textState = textFragment.TextState;
        FontHelper.Pdf.ApplyFontSettings(textState, fontName, fontSize);
        textState.ForegroundColor = Color.FromRgb(ColorHelper.ParseColor(color));

        var textBuilder = new TextBuilder(page);
        textBuilder.AppendText(textFragment);

        MarkModified(context);

        return Success($"Text added to page {pageIndex}.");
    }
}
