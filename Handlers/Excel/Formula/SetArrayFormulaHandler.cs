using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for setting array formulas in Excel.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetArrayFormulaHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_array";

    /// <summary>
    ///     Sets an array formula for a range of cells.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range, formula
    ///     Optional: sheetIndex (default: 0), autoCalculate (default: true)
    /// </param>
    /// <returns>Success message with range details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var setParams = ExtractSetArrayParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, setParams.SheetIndex);

        var rangeObj = ExcelHelper.CreateRange(worksheet.Cells, setParams.Range);
        var cleanFormula = setParams.Formula.TrimStart('{').TrimEnd('}');

        ValidateRange(rangeObj);

        var formulaContext = new FormulaContext(worksheet, rangeObj, cleanFormula, setParams.Range);
        var result = TrySetArrayFormula(formulaContext);

        if (result.IsSuccess)
        {
            if (setParams.AutoCalculate) workbook.CalculateFormula();
            MarkModified(context);
            return new SuccessResult { Message = result.Message };
        }

        throw new ArgumentException(
            $"Failed to set array formula. Range: {setParams.Range}, Formula: {cleanFormula}. Errors: {result.Message}");
    }

    /// <summary>
    ///     Validates the range dimensions and position.
    /// </summary>
    /// <param name="rangeObj">The range object to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the range is invalid.</exception>
    private static void ValidateRange(Aspose.Cells.Range rangeObj)
    {
        if (rangeObj.RowCount <= 0 || rangeObj.ColumnCount <= 0)
            throw new ArgumentException(
                $"Invalid range dimensions: rows={rangeObj.RowCount}, columns={rangeObj.ColumnCount}");

        if (rangeObj.FirstRow < 0 || rangeObj.FirstColumn < 0)
            throw new ArgumentException(
                $"Invalid range position: startRow={rangeObj.FirstRow}, startColumn={rangeObj.FirstColumn}");
    }

    /// <summary>
    ///     Extracts set array parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set array parameters.</returns>
    private static SetArrayParameters ExtractSetArrayParameters(OperationParameters parameters)
    {
        return new SetArrayParameters(
            parameters.GetRequired<string>("range"),
            parameters.GetRequired<string>("formula"),
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("autoCalculate", true)
        );
    }

    /// <summary>
    ///     Record to hold set array formula parameters.
    /// </summary>
    private sealed record SetArrayParameters(string Range, string Formula, int SheetIndex, bool AutoCalculate);

    /// <summary>
    ///     Context for formula operations containing all required data.
    /// </summary>
    /// <param name="Worksheet">The worksheet to operate on.</param>
    /// <param name="RangeObj">The range object.</param>
    /// <param name="CleanFormula">The cleaned formula string.</param>
    /// <param name="Range">The range string representation.</param>
    private sealed record FormulaContext(
        Worksheet Worksheet,
        Aspose.Cells.Range RangeObj,
        string CleanFormula,
        string Range);

    /// <summary>
    ///     Result of a formula operation attempt.
    /// </summary>
    /// <param name="IsSuccess">Whether the operation succeeded.</param>
    /// <param name="Message">The result message or error description.</param>
    private sealed record FormulaResult(bool IsSuccess, string Message);

#pragma warning disable CS0618
    /// <summary>
    ///     Tries to set an array formula using multiple methods.
    /// </summary>
    /// <param name="ctx">The formula context.</param>
    /// <returns>The result of the operation.</returns>
    private static FormulaResult TrySetArrayFormula(FormulaContext ctx)
    {
        List<string> errors = [];

        var result = TryPrimaryMethod(ctx);
        if (result.IsSuccess) return result;
        errors.Add(result.Message);

        result = TryAlternativeMethod(ctx);
        if (result.IsSuccess) return result;
        errors.Add(result.Message);

        result = TryFallbackMethod(ctx);
        if (result.IsSuccess) return result;
        errors.Add(result.Message);

        return new FormulaResult(false, string.Join(", ", errors));
    }

    /// <summary>
    ///     Tries to set array formula using the primary method with SetArrayFormula.
    /// </summary>
    /// <param name="ctx">The formula context.</param>
    /// <returns>The result of the operation.</returns>
    private static FormulaResult TryPrimaryMethod(FormulaContext ctx)
    {
        try
        {
            ClearRange(ctx.Worksheet, ctx.RangeObj);
            var formulaToSet = ctx.CleanFormula.StartsWith('=') ? ctx.CleanFormula : "=" + ctx.CleanFormula;
            var firstCell = ctx.Worksheet.Cells[ctx.RangeObj.FirstRow, ctx.RangeObj.FirstColumn];
            firstCell.SetArrayFormula(formulaToSet, ctx.RangeObj.RowCount, ctx.RangeObj.ColumnCount);
            return new FormulaResult(true, $"Array formula set in range {ctx.Range}.");
        }
        catch (Exception ex)
        {
            return new FormulaResult(false, ex.Message);
        }
    }

    /// <summary>
    ///     Tries to set array formula using an alternative method with 5 parameters.
    /// </summary>
    /// <param name="ctx">The formula context.</param>
    /// <returns>The result of the operation.</returns>
    private static FormulaResult TryAlternativeMethod(FormulaContext ctx)
    {
        try
        {
            ClearRange(ctx.Worksheet, ctx.RangeObj);
            var formulaWithoutEquals = ctx.CleanFormula.StartsWith('=') ? ctx.CleanFormula[1..] : ctx.CleanFormula;
            var firstCell = ctx.Worksheet.Cells[ctx.RangeObj.FirstRow, ctx.RangeObj.FirstColumn];
            firstCell.SetArrayFormula(formulaWithoutEquals, ctx.RangeObj.FirstRow, ctx.RangeObj.FirstColumn, false,
                false);

            if (firstCell.IsArrayFormula)
                return new FormulaResult(true, $"Array formula set in range {ctx.Range}.");

            return new FormulaResult(false, "SetArrayFormula with 5 parameters did not work");
        }
        catch (Exception ex)
        {
            return new FormulaResult(false, ex.Message);
        }
    }

    /// <summary>
    ///     Tries to set formula using the fallback method by setting formula to each cell individually.
    /// </summary>
    /// <param name="ctx">The formula context.</param>
    /// <returns>The result of the operation.</returns>
    private static FormulaResult TryFallbackMethod(FormulaContext ctx)
    {
        try
        {
            var formulaWithEquals = ctx.CleanFormula.StartsWith('=') ? ctx.CleanFormula : "=" + ctx.CleanFormula;
            for (var i = 0; i < ctx.RangeObj.RowCount; i++)
            for (var j = 0; j < ctx.RangeObj.ColumnCount; j++)
                ctx.Worksheet.Cells[ctx.RangeObj.FirstRow + i, ctx.RangeObj.FirstColumn + j].Formula =
                    formulaWithEquals;

            return new FormulaResult(true, $"Formula set to range {ctx.Range} (not a true array formula).");
        }
        catch (Exception ex)
        {
            return new FormulaResult(false, ex.Message);
        }
    }

    /// <summary>
    ///     Clears all cells in the specified range.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the range.</param>
    /// <param name="rangeObj">The range to clear.</param>
    private static void ClearRange(Worksheet worksheet, Aspose.Cells.Range rangeObj)
    {
        for (var i = 0; i < rangeObj.RowCount; i++)
        for (var j = 0; j < rangeObj.ColumnCount; j++)
            worksheet.Cells[rangeObj.FirstRow + i, rangeObj.FirstColumn + j].PutValue("");
    }
#pragma warning restore CS0618
}
