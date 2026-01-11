using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for linearizing a PDF document for fast web viewing.
/// </summary>
public class LinearizePdfFileHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "linearize";

    /// <summary>
    ///     Linearizes a PDF document for fast web viewing.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No additional parameters required.</param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;
        document.Optimize();

        MarkModified(context);

        return Success("PDF linearized for fast web view.");
    }
}
