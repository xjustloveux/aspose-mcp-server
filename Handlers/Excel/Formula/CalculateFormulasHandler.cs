using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for calculating all formulas in an Excel workbook.
/// </summary>
public class CalculateFormulasHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "calculate";

    /// <summary>
    ///     Calculates all formulas in the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        context.Document.CalculateFormula();
        MarkModified(context);

        return Success("Formulas calculated.");
    }
}
