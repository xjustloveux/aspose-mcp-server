using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for removing a field from a pivot table.
/// </summary>
public class DeleteFieldExcelPivotTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete_field";

    /// <summary>
    ///     Removes a field from the pivot table.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: pivotTableIndex, fieldName, fieldType
    ///     Optional: sheetIndex
    /// </param>
    /// <returns>Success message with delete field details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var pivotTableIndex = parameters.GetRequired<int>("pivotTableIndex");
        var fieldName = parameters.GetRequired<string>("fieldName");
        var fieldType = parameters.GetRequired<string>("fieldType");

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

            var fieldTypeEnum = ExcelPivotTableHelper.ParseFieldType(fieldType);

            try
            {
                pivotTable.RemoveField(fieldTypeEnum, fieldIndex);

                try
                {
                    pivotTable.CalculateData();
                }
                catch (Exception calcEx)
                {
                    Console.Error.WriteLine($"[WARN] CalculateData warning: {calcEx.Message}");
                }

                MarkModified(context);

                return Success($"Field '{fieldName}' removed from {fieldType} area of pivot table #{pivotTableIndex}.");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("not found") || ex.Message.Contains("does not exist"))
                {
                    MarkModified(context);
                    return Success(
                        $"Field '{fieldName}' may already be removed from {fieldType} area of pivot table #{pivotTableIndex}.");
                }

                throw new ArgumentException(
                    $"Failed to remove field '{fieldName}' from pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {fieldType}");
            }
        }
        catch (Exception outerEx)
        {
            throw new ArgumentException(
                $"Failed to remove field '{fieldName}' from pivot table: {outerEx.Message}");
        }
    }
}
