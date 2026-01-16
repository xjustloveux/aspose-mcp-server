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
        var p = ExtractRotateParameters(parameters);

        if (p.Rotation != 0 && p.Rotation != 90 && p.Rotation != 180 && p.Rotation != 270)
            throw new ArgumentException("rotation must be 0, 90, 180, or 270");

        var doc = context.Document;

        var rotationEnum = p.Rotation switch
        {
            90 => Rotation.on90,
            180 => Rotation.on180,
            270 => Rotation.on270,
            _ => Rotation.None
        };

        List<int> pagesToRotate;
        if (p.PageIndices is { Length: > 0 })
            pagesToRotate = p.PageIndices.ToList();
        else if (p.PageIndex is > 0)
            pagesToRotate = [p.PageIndex.Value];
        else
            pagesToRotate = Enumerable.Range(1, doc.Pages.Count).ToList();

        foreach (var pageNum in pagesToRotate)
            if (pageNum >= 1 && pageNum <= doc.Pages.Count)
                doc.Pages[pageNum].Rotate = rotationEnum;

        MarkModified(context);

        return Success($"Rotated {pagesToRotate.Count} page(s) by {p.Rotation} degrees.");
    }

    /// <summary>
    ///     Extracts parameters for rotate operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static RotateParameters ExtractRotateParameters(OperationParameters parameters)
    {
        return new RotateParameters(
            parameters.GetRequired<int>("rotation"),
            parameters.GetOptional<int?>("pageIndex"),
            parameters.GetOptional<int[]?>("pageIndices")
        );
    }

    /// <summary>
    ///     Parameters for rotate operation.
    /// </summary>
    /// <param name="Rotation">The rotation angle (0, 90, 180, or 270 degrees).</param>
    /// <param name="PageIndex">The optional 1-based page index.</param>
    /// <param name="PageIndices">The optional array of 1-based page indices.</param>
    private sealed record RotateParameters(int Rotation, int? PageIndex, int[]? PageIndices);
}
