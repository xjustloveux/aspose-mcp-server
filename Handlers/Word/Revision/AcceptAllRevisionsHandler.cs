using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Revision;

/// <summary>
///     Handler for accepting all revisions in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AcceptAllRevisionsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "accept_all";

    /// <summary>
    ///     Accepts all revisions in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>Success message with revision count.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var count = doc.Revisions.Count;
        doc.AcceptAllRevisions();

        MarkModified(context);

        return new SuccessResult { Message = $"Accepted {count} revision(s)" };
    }
}
