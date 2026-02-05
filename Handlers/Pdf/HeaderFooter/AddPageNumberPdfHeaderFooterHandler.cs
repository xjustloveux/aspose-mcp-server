using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Pdf;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.HeaderFooter;

/// <summary>
///     Handler for adding page number stamps to PDF document headers or footers.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddPageNumberPdfHeaderFooterHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_page_number";

    /// <summary>
    ///     Adds page number stamps to the header or footer of specified pages.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: format (default: "Page {0} of {1}"), position (default: footer),
    ///     alignment (default: center), fontSize (default: 10.0), margin (default: 20.0),
    ///     startPage (default: 1), pageRange.
    /// </param>
    /// <returns>Success message with the number of pages affected.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        var document = context.Document;

        var pages = PdfWatermarkHelper.ParsePageRange(p.PageRange, document.Pages.Count);
        var pageNumber = p.StartPage;

        foreach (var pageIdx in pages)
        {
            var pageText = string.Format(p.Format, pageNumber, document.Pages.Count);
            var stamp = new TextStamp(pageText)
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

            stamp.HorizontalAlignment = AddTextPdfHeaderFooterHandler.ResolveHorizontalAlignment(p.Alignment);

            document.Pages[pageIdx].AddStamp(stamp);
            pageNumber++;
        }

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Added page numbers to {pages.Count} page(s)."
        };
    }

    /// <summary>
    ///     Extracts add page number parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddPageNumberParameters ExtractParameters(OperationParameters parameters)
    {
        return new AddPageNumberParameters(
            parameters.GetOptional("format", "Page {0} of {1}"),
            parameters.GetOptional("position", "footer"),
            parameters.GetOptional("alignment", "center"),
            parameters.GetOptional("fontSize", 10.0),
            parameters.GetOptional("margin", 20.0),
            parameters.GetOptional("startPage", 1),
            parameters.GetOptional<string?>("pageRange")
        );
    }

    /// <summary>
    ///     Parameters for adding page numbers to headers or footers.
    /// </summary>
    /// <param name="Format">The page number format string ({0} = current page, {1} = total pages).</param>
    /// <param name="Position">The position (header or footer).</param>
    /// <param name="Alignment">The horizontal alignment (left, center, or right).</param>
    /// <param name="FontSize">The font size in points.</param>
    /// <param name="Margin">The margin from the edge in points.</param>
    /// <param name="StartPage">The starting page number.</param>
    /// <param name="PageRange">The optional page range string.</param>
    private sealed record AddPageNumberParameters(
        string Format,
        string Position,
        string Alignment,
        double FontSize,
        double Margin,
        int StartPage,
        string? PageRange);
}
