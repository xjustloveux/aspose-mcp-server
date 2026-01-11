using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Watermark;

/// <summary>
///     Handler for removing watermarks from Word documents.
/// </summary>
public class RemoveWatermarkWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes watermark from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>Success message indicating watermark removal status.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;

        if (doc.Watermark.Type == WatermarkType.None)
            return Success("No watermark found in document");

        doc.Watermark.Remove();

        MarkModified(context);

        return Success("Watermark removed from document");
    }
}
