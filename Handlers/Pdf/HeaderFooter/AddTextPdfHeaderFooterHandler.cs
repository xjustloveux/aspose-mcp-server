using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Pdf;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.HeaderFooter;

/// <summary>
///     Handler for adding text stamps to PDF document headers or footers.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddTextPdfHeaderFooterHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_text";

    /// <summary>
    ///     Adds a text stamp to the header or footer of specified pages.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text.
    ///     Optional: position (default: header), alignment (default: center),
    ///     fontSize (default: 12.0), margin (default: 20.0), pageRange.
    /// </param>
    /// <returns>Success message with the number of pages affected.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        var document = context.Document;

        var stamp = new TextStamp(p.Text)
        {
            TextState = { FontSize = (float)p.FontSize, FontStyle = FontStyles.Regular }
        };

        if (p.Position.Equals("header", StringComparison.OrdinalIgnoreCase))
        {
            stamp.VerticalAlignment = VerticalAlignment.Top;
            stamp.TopMargin = p.Margin;
        }
        else
        {
            stamp.VerticalAlignment = VerticalAlignment.Bottom;
            stamp.BottomMargin = p.Margin;
        }

        stamp.HorizontalAlignment = ResolveHorizontalAlignment(p.Alignment);

        var pages = PdfWatermarkHelper.ParsePageRange(p.PageRange, document.Pages.Count);
        foreach (var pageIdx in pages) document.Pages[pageIdx].AddStamp(stamp);

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Added text {p.Position} to {pages.Count} page(s)."
        };
    }

    /// <summary>
    ///     Resolves a string alignment name to the corresponding <see cref="HorizontalAlignment" /> enum value.
    /// </summary>
    /// <param name="alignment">The alignment name (left, center, or right).</param>
    /// <returns>The corresponding <see cref="HorizontalAlignment" /> value.</returns>
    internal static HorizontalAlignment ResolveHorizontalAlignment(string alignment)
    {
        return alignment.ToLowerInvariant() switch
        {
            "left" => HorizontalAlignment.Left,
            "right" => HorizontalAlignment.Right,
            _ => HorizontalAlignment.Center
        };
    }

    /// <summary>
    ///     Extracts add text header/footer parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddTextParameters ExtractParameters(OperationParameters parameters)
    {
        return new AddTextParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional("position", "header"),
            parameters.GetOptional("alignment", "center"),
            parameters.GetOptional("fontSize", 12.0),
            parameters.GetOptional("margin", 20.0),
            parameters.GetOptional<string?>("pageRange")
        );
    }

    /// <summary>
    ///     Parameters for adding a text header or footer.
    /// </summary>
    /// <param name="Text">The text content to add.</param>
    /// <param name="Position">The position (header or footer).</param>
    /// <param name="Alignment">The horizontal alignment (left, center, or right).</param>
    /// <param name="FontSize">The font size in points.</param>
    /// <param name="Margin">The margin from the edge in points.</param>
    /// <param name="PageRange">The optional page range string.</param>
    private sealed record AddTextParameters(
        string Text,
        string Position,
        string Alignment,
        double FontSize,
        double Margin,
        string? PageRange);
}
