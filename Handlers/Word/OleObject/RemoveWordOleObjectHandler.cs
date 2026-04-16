using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.Word.OleObject;

/// <summary>
///     Handler for the <c>remove</c> operation on <c>word_ole_object</c>. Removes the
///     selected OLE-bearing shape from the document and marks the context modified so
///     the tool layer persists the change (file-mode re-save or session-mode mark-dirty).
/// </summary>
[ResultType(typeof(OleRemoveResult))]
public sealed class RemoveWordOleObjectHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes an OLE-bearing shape by flat index.
    /// </summary>
    /// <param name="context">Operation context; <c>Document</c> must be non-null.</param>
    /// <param name="parameters">Required: <c>oleIndex</c>.</param>
    /// <returns>An <see cref="OleRemoveResult" /> confirming removal.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the index is out of range.</exception>
    /// <exception cref="IOException">Thrown when Aspose rejects the remove call.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(parameters);

        var oleIndex = parameters.GetRequired<int>(OleParamKeys.OleIndex);
        var (shape, flatIndex) = OleHandlerShared.LocateWordShape(context.Document, oleIndex);

        try
        {
            shape.Remove();
        }
        catch (Exception ex)
        {
            throw OleErrorTranslator.Translate(ex, Path.GetFileName(context.SourcePath));
        }

        MarkModified(context);

        return new OleRemoveResult
        {
            Index = flatIndex,
            Removed = true,
            SavedTo = context.SessionId == null ? context.OutputPath ?? context.SourcePath : null
        };
    }
}
