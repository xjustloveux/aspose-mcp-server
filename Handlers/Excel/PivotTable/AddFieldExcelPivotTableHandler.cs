using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for adding a field to a pivot table.
/// </summary>
public class AddFieldExcelPivotTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add_field";

    /// <summary>
    ///     Adds a field to the pivot table.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: pivotTableIndex, fieldName, fieldType
    ///     Optional: sheetIndex, function
    /// </param>
    /// <returns>Success message with add field details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var pivotTableIndex = parameters.GetRequired<int>("pivotTableIndex");
        var fieldName = parameters.GetRequired<string>("fieldName");
        var fieldType = parameters.GetRequired<string>("fieldType");
        var function = parameters.GetOptional("function", "Sum");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pivotTables = worksheet.PivotTables;

            if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
                throw new ArgumentException(
                    $"Pivot table index {pivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

            var pivotTable = pivotTables[pivotTableIndex];

            var (sourceSheet, sourceRangeObj) = ExcelPivotTableHelper.ParseDataSource(
                workbook, pivotTable, sheetIndex, pivotTableIndex, worksheet.Name);

            var fieldIndex = ExcelPivotTableHelper.FindFieldIndex(sourceSheet, sourceRangeObj, fieldName);

            if (fieldIndex < 0)
            {
                var availableFields = ExcelPivotTableHelper.GetAvailableFieldNames(sourceSheet, sourceRangeObj);
                var availableFieldsStr = availableFields.Count > 0
                    ? $" Available fields in header row: {string.Join(", ", availableFields)}"
                    : " No field names found in header row.";

                throw new ArgumentException(
                    $"Field '{fieldName}' not found in pivot table source data.{availableFieldsStr} Please check that the field name matches exactly (case-sensitive).");
            }

            try
            {
                var pivotFieldType = ExcelPivotTableHelper.ParseFieldType(fieldType);
                pivotTable.AddFieldToArea(pivotFieldType, fieldIndex);

                if (pivotFieldType == PivotFieldType.Data && pivotTable.DataFields.Count > 0)
                {
                    var dataField = pivotTable.DataFields[^1];
                    dataField.Function = ExcelPivotTableHelper.ParseFunction(function);
                }

                try
                {
                    pivotTable.CalculateData();
                }
                catch (Exception calcEx)
                {
                    Console.Error.WriteLine($"[WARN] CalculateData warning: {calcEx.Message}");
                }

                MarkModified(context);

                return Success($"Field '{fieldName}' added as {fieldType} field to pivot table #{pivotTableIndex}.");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("already exists") || ex.Message.Contains("duplicate"))
                {
                    MarkModified(context);
                    return Success(
                        $"Field '{fieldName}' may already exist in {fieldType} area of pivot table #{pivotTableIndex}.");
                }

                throw new ArgumentException(
                    $"Failed to add field '{fieldName}' to pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {fieldType}");
            }
        }
        catch (Exception outerEx)
        {
            throw new ArgumentException(
                $"Failed to add field '{fieldName}' to pivot table: {outerEx.Message}");
        }
    }
}
