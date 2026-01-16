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
        var p = ExtractDeleteFieldParameters(parameters);

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

                return Success(
                    $"Field '{p.FieldName}' removed from {p.FieldType} area of pivot table #{p.PivotTableIndex}.");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("not found") || ex.Message.Contains("does not exist"))
                {
                    MarkModified(context);
                    return Success(
                        $"Field '{p.FieldName}' may already be removed from {p.FieldType} area of pivot table #{p.PivotTableIndex}.");
                }

                throw new ArgumentException(
                    $"Failed to remove field '{p.FieldName}' from pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {p.FieldType}");
            }
        }
        catch (Exception outerEx)
        {
            throw new ArgumentException(
                $"Failed to remove field '{p.FieldName}' from pivot table: {outerEx.Message}");
        }
    }

    private static DeleteFieldParameters ExtractDeleteFieldParameters(OperationParameters parameters)
    {
        return new DeleteFieldParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("pivotTableIndex"),
            parameters.GetRequired<string>("fieldName"),
            parameters.GetRequired<string>("fieldType")
        );
    }

    private record DeleteFieldParameters(int SheetIndex, int PivotTableIndex, string FieldName, string FieldType);
}
