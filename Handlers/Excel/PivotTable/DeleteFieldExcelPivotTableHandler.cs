using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for removing a field from a pivot table.
/// </summary>
public class DeleteFieldExcelPivotTableHandler : ExcelPivotTableFieldHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "delete_field";

    /// <inheritdoc />
    protected override string OperationVerb => "remove";

    /// <inheritdoc />
    protected override string OperationVerbPast => "removed";

    /// <inheritdoc />
    protected override string GetPreposition()
    {
        return "from";
    }

    /// <summary>
    ///     Extracts field parameters from operation parameters.
    ///     Overridden to exclude the function parameter which is not needed for delete.
    /// </summary>
    /// <param name="parameters">The operation parameters to extract from.</param>
    /// <returns>The extracted <see cref="ExcelPivotTableFieldHandlerBase.FieldParameters" />.</returns>
    protected override FieldParameters ExtractFieldParameters(OperationParameters parameters)
    {
        return new FieldParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("pivotTableIndex"),
            parameters.GetRequired<string>("fieldName"),
            parameters.GetRequired<string>("fieldType"),
            "Sum"
        );
    }

    /// <summary>
    ///     Removes a field from the specified area of the pivot table.
    /// </summary>
    /// <param name="context">The workbook context containing the Excel workbook.</param>
    /// <param name="pivotTable">The target pivot table.</param>
    /// <param name="parameters">The extracted field parameters.</param>
    /// <param name="fieldIndex">The zero-based field index in the source data.</param>
    /// <param name="fieldType">The pivot field type (Row, Column, Page, Data).</param>
    /// <returns>A <see cref="SuccessResult" /> indicating the field was removed.</returns>
    /// <exception cref="ArgumentException">Thrown when removing the field fails.</exception>
    protected override SuccessResult ExecuteFieldOperation(
        OperationContext<Workbook> context,
        Aspose.Cells.Pivot.PivotTable pivotTable,
        FieldParameters parameters,
        int fieldIndex,
        PivotFieldType fieldType)
    {
        try
        {
            pivotTable.RemoveField(fieldType, fieldIndex);

            CalculatePivotTableData(pivotTable);
            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Field '{parameters.FieldName}' removed from {parameters.FieldType} area of pivot table #{parameters.PivotTableIndex}."
            };
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("not found") || ex.Message.Contains("does not exist"))
            {
                MarkModified(context);
                return new SuccessResult
                {
                    Message =
                        $"Field '{parameters.FieldName}' may already be removed from {parameters.FieldType} area of pivot table #{parameters.PivotTableIndex}."
                };
            }

            throw new ArgumentException(
                $"Failed to remove field '{parameters.FieldName}' from pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {parameters.FieldType}");
        }
    }
}
