using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Stamp;

/// <summary>
///     Handler for adding PDF page stamps to PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddPdfStampHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_pdf";

    /// <summary>
    ///     Adds a PDF page stamp from a source PDF to one or all pages in the target PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pdfPath.
    ///     Optional: stampPageIndex (default: 1), pageIndex (default: 0, 0 = all pages), x, y, width, height, opacity,
    ///     rotation.
    /// </param>
    /// <returns>Success message with stamp details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or page index is out of range.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the source PDF file is not found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        if (string.IsNullOrEmpty(p.PdfPath))
            throw new ArgumentException("pdfPath is required for add_pdf operation");

        SecurityHelper.ValidateFilePath(p.PdfPath, "pdfPath", true);

        if (!File.Exists(p.PdfPath))
            throw new FileNotFoundException($"PDF file not found: {p.PdfPath}");

        var document = context.Document;

        var stamp = new PdfPageStamp(p.PdfPath, p.StampPageIndex)
        {
            Opacity = p.Opacity,
            RotateAngle = p.Rotation
        };
        if (p.X > 0) stamp.XIndent = p.X;
        if (p.Y > 0) stamp.YIndent = p.Y;
        if (p.Width > 0) stamp.Width = p.Width;
        if (p.Height > 0) stamp.Height = p.Height;

        if (p.PageIndex == 0)
        {
            for (var i = 1; i <= document.Pages.Count; i++)
                document.Pages[i].AddStamp(stamp);
        }
        else
        {
            AddTextPdfStampHandler.ValidatePageIndex(p.PageIndex, document);
            document.Pages[p.PageIndex].AddStamp(stamp);
        }

        MarkModified(context);

        var pageDesc = p.PageIndex == 0 ? "all pages" : $"page {p.PageIndex}";
        return new SuccessResult
        {
            Message =
                $"PDF page stamp added to {pageDesc}. Source: {Path.GetFileName(p.PdfPath)}, stamp page: {p.StampPageIndex}"
        };
    }

    /// <summary>
    ///     Extracts parameters for the add PDF stamp operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddPdfStampParameters ExtractParameters(OperationParameters parameters)
    {
        return new AddPdfStampParameters(
            parameters.GetRequired<string>("pdfPath"),
            parameters.GetOptional("stampPageIndex", 1),
            parameters.GetOptional("pageIndex", 0),
            parameters.GetOptional("x", 0.0),
            parameters.GetOptional("y", 0.0),
            parameters.GetOptional("width", 0.0),
            parameters.GetOptional("height", 0.0),
            parameters.GetOptional("opacity", 1.0),
            parameters.GetOptional("rotation", 0.0)
        );
    }

    /// <summary>
    ///     Parameters for the add PDF stamp operation.
    /// </summary>
    /// <param name="PdfPath">The path to the source PDF file.</param>
    /// <param name="StampPageIndex">The 1-based page index from the source PDF to use as stamp.</param>
    /// <param name="PageIndex">The target page index (0 = all pages, otherwise 1-based).</param>
    /// <param name="X">The X position indent.</param>
    /// <param name="Y">The Y position indent.</param>
    /// <param name="Width">The stamp width (0 = auto).</param>
    /// <param name="Height">The stamp height (0 = auto).</param>
    /// <param name="Opacity">The stamp opacity (0.0 to 1.0).</param>
    /// <param name="Rotation">The rotation angle in degrees.</param>
    private sealed record AddPdfStampParameters(
        string PdfPath,
        int StampPageIndex,
        int PageIndex,
        double X,
        double Y,
        double Width,
        double Height,
        double Opacity,
        double Rotation);
}
