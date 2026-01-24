using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for adding a field to a pivot table.
/// </summary>
public class AddFieldExcelPivotTableHandler : ExcelPivotTableFieldHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "add_field";

    /// <inheritdoc />
    protected override string OperationVerb => "add";

    /// <inheritdoc />
    protected override string OperationVerbPast => "added";

    /// <inheritdoc />
    protected override string GetPreposition()
    {
        return "to";
    }

    /// <summary>
    ///     Adds a field to the specified area of the pivot table.
    /// </summary>
    /// <param name="context">The workbook context containing the Excel workbook.</param>
    /// <param name="pivotTable">The target pivot table.</param>
    /// <param name="parameters">The extracted field parameters.</param>
    /// <param name="fieldIndex">The zero-based field index in the source data.</param>
    /// <param name="fieldType">The pivot field type (Row, Column, Page, Data).</param>
    /// <returns>A <see cref="SuccessResult" /> indicating the field was added.</returns>
    /// <exception cref="ArgumentException">Thrown when adding the field fails.</exception>
    protected override SuccessResult ExecuteFieldOperation(
        OperationContext<Workbook> context,
        Aspose.Cells.Pivot.PivotTable pivotTable,
        FieldParameters parameters,
        int fieldIndex,
        PivotFieldType fieldType)
    {
        try
        {
            pivotTable.AddFieldToArea(fieldType, fieldIndex);

            if (fieldType == PivotFieldType.Data && pivotTable.DataFields.Count > 0)
            {
                var dataField = pivotTable.DataFields[^1];
                dataField.Function = ExcelPivotTableHelper.ParseFunction(parameters.Function);
            }

            CalculatePivotTableData(pivotTable);
            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Field '{parameters.FieldName}' added as {parameters.FieldType} field to pivot table #{parameters.PivotTableIndex}."
            };
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("already exists") || ex.Message.Contains("duplicate"))
            {
                MarkModified(context);
                return new SuccessResult
                {
                    Message =
                        $"Field '{parameters.FieldName}' may already exist in {parameters.FieldType} area of pivot table #{parameters.PivotTableIndex}."
                };
            }

            throw new ArgumentException(
                $"Failed to add field '{parameters.FieldName}' to pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {parameters.FieldType}");
        }
    }
}
