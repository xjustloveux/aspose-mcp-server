using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Watermark;

/// <summary>
///     Handler for removing watermarks from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;

        if (doc.Watermark.Type == WatermarkType.None)
            return new SuccessResult { Message = "No watermark found in document" };

        doc.Watermark.Remove();

        MarkModified(context);

        return new SuccessResult { Message = "Watermark removed from document" };
    }
}
