using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Stamp;

/// <summary>
///     Handler for adding text stamps to PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddTextPdfStampHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_text";

    /// <summary>
    ///     Adds a text stamp to one or all pages in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text.
    ///     Optional: pageIndex (default: 0, 0 = all pages), x, y, fontSize, opacity, rotation, color.
    /// </param>
    /// <returns>Success message with stamp details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or page index is out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        if (string.IsNullOrEmpty(p.Text))
            throw new ArgumentException("text is required for add_text operation");

        var document = context.Document;

        var stamp = new TextStamp(p.Text)
        {
            TextState = { FontSize = (float)p.FontSize },
            Opacity = p.Opacity,
            RotateAngle = p.Rotation
        };

        if (p.X > 0) stamp.XIndent = p.X;
        if (p.Y > 0) stamp.YIndent = p.Y;

        if (!string.IsNullOrEmpty(p.Color))
            stamp.TextState.ForegroundColor = Color.FromRgb(ColorHelper.ParseColor(p.Color));

        if (p.PageIndex == 0)
        {
            for (var i = 1; i <= document.Pages.Count; i++)
                document.Pages[i].AddStamp(stamp);
        }
        else
        {
            ValidatePageIndex(p.PageIndex, document);
            document.Pages[p.PageIndex].AddStamp(stamp);
        }

        MarkModified(context);

        var pageDesc = p.PageIndex == 0 ? "all pages" : $"page {p.PageIndex}";
        return new SuccessResult { Message = $"Text stamp added to {pageDesc}." };
    }

    /// <summary>
    ///     Validates that a page index is within the valid range for the document.
    /// </summary>
    /// <param name="pageIndex">The 1-based page index to validate.</param>
    /// <param name="document">The PDF document.</param>
    /// <exception cref="ArgumentException">Thrown when the page index is out of range.</exception>
    internal static void ValidatePageIndex(int pageIndex, Document document)
    {
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
    }

    /// <summary>
    ///     Extracts parameters for the add text stamp operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddTextStampParameters ExtractParameters(OperationParameters parameters)
    {
        return new AddTextStampParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetOptional("pageIndex", 0),
            parameters.GetOptional("x", 0.0),
            parameters.GetOptional("y", 0.0),
            parameters.GetOptional("fontSize", 14.0),
            parameters.GetOptional("opacity", 1.0),
            parameters.GetOptional("rotation", 0.0),
            parameters.GetOptional("color", "black")
        );
    }

    /// <summary>
    ///     Parameters for the add text stamp operation.
    /// </summary>
    /// <param name="Text">The stamp text content.</param>
    /// <param name="PageIndex">The target page index (0 = all pages, otherwise 1-based).</param>
    /// <param name="X">The X position indent.</param>
    /// <param name="Y">The Y position indent.</param>
    /// <param name="FontSize">The font size for the stamp text.</param>
    /// <param name="Opacity">The stamp opacity (0.0 to 1.0).</param>
    /// <param name="Rotation">The rotation angle in degrees.</param>
    /// <param name="Color">The text color name or hex value.</param>
    private sealed record AddTextStampParameters(
        string Text,
        int PageIndex,
        double X,
        double Y,
        double FontSize,
        double Opacity,
        double Rotation,
        string Color);
}
