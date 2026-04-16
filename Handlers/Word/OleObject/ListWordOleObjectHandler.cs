using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.Word.OleObject;

/// <summary>
///     Handler for the <c>list</c> operation on <c>word_ole_object</c>. Enumerates all
///     OLE-bearing shapes in a <see cref="Document" /> and projects each one to the
///     cross-tool <see cref="OleObjectMetadata" /> shape.
/// </summary>
[ResultType(typeof(OleListResult))]
public sealed class ListWordOleObjectHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "list";

    /// <summary>
    ///     Executes the <c>list</c> operation.
    /// </summary>
    /// <param name="context">Operation context; <c>Document</c> must be non-null.</param>
    /// <param name="parameters">Unused for <c>list</c>.</param>
    /// <returns>An <see cref="OleListResult" /> enumerating every OLE shape.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="context" /> is null.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        var document = context.Document;

        var items = new List<OleObjectMetadata>();
        var flatIndex = 0;
        foreach (var shape in document.GetChildNodes(NodeType.Shape, true).OfType<Aspose.Words.Drawing.Shape>())
        {
            if (shape.OleFormat == null) continue;

            var size = WordOleMetadataMapper.ComputeSize(shape);
            var location = WordOleMetadataMapper.ResolveLocation(document, shape);
            items.Add(WordOleMetadataMapper.Map(shape, flatIndex, location, size));
            flatIndex++;
        }

        return new OleListResult { Count = items.Count, Items = items };
    }
}
