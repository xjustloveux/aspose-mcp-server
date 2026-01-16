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
        var p = ExtractAddFieldParameters(parameters);

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

            try
            {
                var pivotFieldType = ExcelPivotTableHelper.ParseFieldType(p.FieldType);
                pivotTable.AddFieldToArea(pivotFieldType, fieldIndex);

                if (pivotFieldType == PivotFieldType.Data && pivotTable.DataFields.Count > 0)
                {
                    var dataField = pivotTable.DataFields[^1];
                    dataField.Function = ExcelPivotTableHelper.ParseFunction(p.Function);
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

                return Success(
                    $"Field '{p.FieldName}' added as {p.FieldType} field to pivot table #{p.PivotTableIndex}.");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("already exists") || ex.Message.Contains("duplicate"))
                {
                    MarkModified(context);
                    return Success(
                        $"Field '{p.FieldName}' may already exist in {p.FieldType} area of pivot table #{p.PivotTableIndex}.");
                }

                throw new ArgumentException(
                    $"Failed to add field '{p.FieldName}' to pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {p.FieldType}");
            }
        }
        catch (Exception outerEx)
        {
            throw new ArgumentException(
                $"Failed to add field '{p.FieldName}' to pivot table: {outerEx.Message}");
        }
    }

    private static AddFieldParameters ExtractAddFieldParameters(OperationParameters parameters)
    {
        return new AddFieldParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("pivotTableIndex"),
            parameters.GetRequired<string>("fieldName"),
            parameters.GetRequired<string>("fieldType"),
            parameters.GetOptional("function", "Sum")
        );
    }

    private sealed record AddFieldParameters(
        int SheetIndex,
        int PivotTableIndex,
        string FieldName,
        string FieldType,
        string Function);
}
