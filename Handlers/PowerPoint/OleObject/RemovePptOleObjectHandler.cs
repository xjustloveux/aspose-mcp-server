using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.PowerPoint.OleObject;

/// <summary>
///     Handler for the <c>remove</c> operation on <c>ppt_ole_object</c>.
/// </summary>
[ResultType(typeof(OleRemoveResult))]
public sealed class RemovePptOleObjectHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>Removes an OLE frame by flat index via <c>ShapeCollection.Remove</c>.</summary>
    /// <param name="context">Operation context.</param>
    /// <param name="parameters">Required: <c>oleIndex</c>.</param>
    /// <returns>An <see cref="OleRemoveResult" /> confirming removal.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the index is out of range.</exception>
    /// <exception cref="IOException">Thrown when the remove call fails.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(parameters);

        var oleIndex = parameters.GetRequired<int>(OleParamKeys.OleIndex);
        var (frame, slide, _, _) = OleHandlerShared.LocatePptFrame(context.Document, oleIndex);

        try
        {
            slide.Shapes.Remove(frame);
        }
        catch (Exception ex)
        {
            throw OleErrorTranslator.Translate(ex, Path.GetFileName(context.SourcePath));
        }

        MarkModified(context);

        return new OleRemoveResult
        {
            Index = oleIndex,
            Removed = true,
            SavedTo = context.SessionId == null ? context.OutputPath ?? context.SourcePath : null
        };
    }
}
