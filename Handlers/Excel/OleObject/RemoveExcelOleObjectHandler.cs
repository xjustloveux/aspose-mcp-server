using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Ole;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.Excel.OleObject;

/// <summary>
///     Handler for the <c>remove</c> operation on <c>excel_ole_object</c>. Uses the sole
///     supported API <c>OleObjectCollection.RemoveAt(int)</c> per api-verification.md item 3.
/// </summary>
[ResultType(typeof(OleRemoveResult))]
public sealed class RemoveExcelOleObjectHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>Removes an OLE object by flat index.</summary>
    /// <param name="context">Operation context.</param>
    /// <param name="parameters">Required: <c>oleIndex</c>.</param>
    /// <returns>An <see cref="OleRemoveResult" /> confirming removal.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the index is out of range.</exception>
    /// <exception cref="IOException">Thrown when the remove call fails.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(parameters);

        var oleIndex = parameters.GetRequired<int>(OleParamKeys.OleIndex);
        var (_, worksheet, _, localIndex) = OleHandlerShared.LocateExcelOle(context.Document, oleIndex);

        try
        {
            worksheet.OleObjects.RemoveAt(localIndex);
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
