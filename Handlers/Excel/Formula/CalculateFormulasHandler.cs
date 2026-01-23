using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for calculating all formulas in an Excel workbook.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        context.Document.CalculateFormula();
        MarkModified(context);

        return new SuccessResult { Message = "Formulas calculated." };
    }
}
