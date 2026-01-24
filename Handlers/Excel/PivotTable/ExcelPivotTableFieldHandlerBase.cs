using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Base handler for pivot table field operations (add/remove fields).
/// </summary>
[ResultType(typeof(SuccessResult))]
public abstract class ExcelPivotTableFieldHandlerBase : OperationHandlerBase<Workbook>
{
    /// <summary>
    ///     Gets the operation verb for messages (e.g., "add", "remove").
    /// </summary>
    protected abstract string OperationVerb { get; }

    /// <summary>
    ///     Gets the operation verb in past tense for messages (e.g., "added", "removed").
    /// </summary>
    protected abstract string OperationVerbPast { get; }

    /// <summary>
    ///     Executes the field operation on the pivot table.
    /// </summary>
    /// <param name="context">The workbook context containing the Excel workbook.</param>
    /// <param name="parameters">
    ///     The operation parameters.
    ///     Required: pivotTableIndex, fieldName, fieldType.
    ///     Optional: sheetIndex, function.
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> with operation details.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when pivotTableIndex is out of range or fieldName is not found.
    /// </exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractFieldParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
            var pivotTables = worksheet.PivotTables;

            if (p.PivotTableIndex < 0 || p.PivotTableIndex >= pivotTables.Count)
                throw new ArgumentException(
                    $"Pivot table index {p.PivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

            var pivotTable = pivotTables[p.PivotTableIndex];

            var (sourceSheet, sourceRangeObj) = ExcelPivotTableHelper.ParseDataSource(
                workbook, pivotTable, p.SheetIndex, p.PivotTableIndex, worksheet.Name);

            var fieldIndex = ExcelPivotTableHelper.FindFieldIndex(sourceSheet, sourceRangeObj, p.FieldName);

            if (fieldIndex < 0)
            {
                var availableFields = ExcelPivotTableHelper.GetAvailableFieldNames(sourceSheet, sourceRangeObj);
                var availableFieldsStr = availableFields.Count > 0
                    ? $" Available fields in header row: {string.Join(", ", availableFields)}"
                    : " No field names found in header row.";

                throw new ArgumentException(
                    $"Field '{p.FieldName}' not found in pivot table source data.{availableFieldsStr} Please check that the field name matches exactly (case-sensitive).");
            }

            var fieldTypeEnum = ExcelPivotTableHelper.ParseFieldType(p.FieldType);

            return ExecuteFieldOperation(context, pivotTable, p, fieldIndex, fieldTypeEnum);
        }
        catch (Exception outerEx)
        {
            throw new ArgumentException(
                $"Failed to {OperationVerb} field '{p.FieldName}' {GetPreposition()} pivot table: {outerEx.Message}");
        }
    }

    /// <summary>
    ///     Executes the specific field operation (add or remove).
    /// </summary>
    /// <param name="context">The workbook context containing the Excel workbook.</param>
    /// <param name="pivotTable">The target pivot table.</param>
    /// <param name="parameters">The extracted field parameters.</param>
    /// <param name="fieldIndex">The zero-based field index in the source data.</param>
    /// <param name="fieldType">The pivot field type (Row, Column, Page, Data).</param>
    /// <returns>A <see cref="SuccessResult" /> containing the operation result.</returns>
    /// <exception cref="ArgumentException">Thrown when the field operation fails.</exception>
    protected abstract SuccessResult ExecuteFieldOperation(
        OperationContext<Workbook> context,
        Aspose.Cells.Pivot.PivotTable pivotTable,
        FieldParameters parameters,
        int fieldIndex,
        PivotFieldType fieldType);

    /// <summary>
    ///     Gets the preposition for the operation message.
    /// </summary>
    /// <returns>The preposition string (e.g., "to" for add, "from" for remove).</returns>
    protected abstract string GetPreposition();

    /// <summary>
    ///     Extracts field parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters to extract from.</param>
    /// <returns>The extracted <see cref="FieldParameters" />.</returns>
    protected virtual FieldParameters ExtractFieldParameters(OperationParameters parameters)
    {
        return new FieldParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("pivotTableIndex"),
            parameters.GetRequired<string>("fieldName"),
            parameters.GetRequired<string>("fieldType"),
            parameters.GetOptional("function", "Sum")
        );
    }

    /// <summary>
    ///     Calculates pivot table data with warning suppression.
    ///     Warnings during calculation are logged but do not throw exceptions.
    /// </summary>
    /// <param name="pivotTable">The pivot table to calculate.</param>
    protected static void CalculatePivotTableData(Aspose.Cells.Pivot.PivotTable pivotTable)
    {
        try
        {
            pivotTable.CalculateData();
        }
        catch (Exception calcEx)
        {
            Console.Error.WriteLine($"[WARN] CalculateData warning: {calcEx.Message}");
        }
    }

    /// <summary>
    ///     Record to hold field operation parameters.
    /// </summary>
    /// <param name="SheetIndex">The zero-based worksheet index (default: 0).</param>
    /// <param name="PivotTableIndex">The zero-based pivot table index.</param>
    /// <param name="FieldName">The field name to add or remove.</param>
    /// <param name="FieldType">The field type (Row, Column, Page, Data).</param>
    /// <param name="Function">The aggregation function for data fields (default: Sum).</param>
    protected sealed record FieldParameters(
        int SheetIndex,
        int PivotTableIndex,
        string FieldName,
        string FieldType,
        string Function);
}
