using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for deleting paragraphs from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteParagraphWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a paragraph at the specified index.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParagraphParameters(parameters);

        if (!deleteParams.ParagraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for delete operation");

        var doc = context.Document;

        var paragraphRef = ParagraphResolver.Resolve(doc,
            ParagraphAddress.From(parameters, deleteParams.ParagraphIndex.Value));
        var idx = paragraphRef.Address.Index;
        var paragraphToDelete = paragraphRef.Paragraph;

        var textPreview = paragraphToDelete.GetText().Trim();
        if (textPreview.Length > 50) textPreview = string.Concat(textPreview.AsSpan(0, 50), "...");

        paragraphToDelete.Remove();

        MarkModified(context);

        var message = $"Paragraph #{idx} deleted successfully\n";
        if (!string.IsNullOrEmpty(textPreview)) message += $"Content preview: {textPreview}\n";
        message += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}";

        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Extracts delete paragraph parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete paragraph parameters.</returns>
    private static DeleteParagraphParameters ExtractDeleteParagraphParameters(OperationParameters parameters)
    {
        return new DeleteParagraphParameters(
            parameters.GetOptional<int?>("paragraphIndex")
        );
    }

    /// <summary>
    ///     Record to hold delete paragraph parameters.
    /// </summary>
    /// <param name="ParagraphIndex">The paragraph index to delete (-1 for last).</param>
    private sealed record DeleteParagraphParameters(int? ParagraphIndex);
}
