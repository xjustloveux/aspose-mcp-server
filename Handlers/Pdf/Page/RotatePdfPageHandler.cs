using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Handler for rotating pages in PDF documents.
/// </summary>
public class RotatePdfPageHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "rotate";

    /// <summary>
    ///     Rotates one or more pages in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: rotation (0, 90, 180, or 270 degrees)
    ///     Optional: pageIndex (1-based page index),
    ///     pageIndices (array of 1-based page indices)
    ///     If neither pageIndex nor pageIndices is provided, all pages are rotated.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when rotation value is invalid.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var rotation = parameters.GetRequired<int>("rotation");
        var pageIndex = parameters.GetOptional<int?>("pageIndex");
        var pageIndices = parameters.GetOptional<int[]?>("pageIndices");

        if (rotation != 0 && rotation != 90 && rotation != 180 && rotation != 270)
            throw new ArgumentException("rotation must be 0, 90, 180, or 270");

        var doc = context.Document;

        var rotationEnum = rotation switch
        {
            90 => Rotation.on90,
            180 => Rotation.on180,
            270 => Rotation.on270,
            _ => Rotation.None
        };

        List<int> pagesToRotate;
        if (pageIndices is { Length: > 0 })
            pagesToRotate = pageIndices.ToList();
        else if (pageIndex is > 0)
            pagesToRotate = [pageIndex.Value];
        else
            pagesToRotate = Enumerable.Range(1, doc.Pages.Count).ToList();

        foreach (var pageNum in pagesToRotate)
            if (pageNum >= 1 && pageNum <= doc.Pages.Count)
                doc.Pages[pageNum].Rotate = rotationEnum;

        MarkModified(context);

        return Success($"Rotated {pagesToRotate.Count} page(s) by {rotation} degrees.");
    }
}
