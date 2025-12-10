using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeMcpServer.Tools;

public class ExcelAddPivotTableFieldTool : IAsposeTool
{
    public string Description => "Add field to pivot table (row, column, data, or page field)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            pivotTableIndex = new
            {
                type = "number",
                description = "Pivot table index (0-based)"
            },
            fieldName = new
            {
                type = "string",
                description = "Field name from source data"
            },
            fieldType = new
            {
                type = "string",
                description = "Field type: 'Row', 'Column', 'Data', 'Page'",
                @enum = new[] { "Row", "Column", "Data", "Page" }
            },
            function = new
            {
                type = "string",
                description = "Aggregation function for data field: 'Sum', 'Count', 'Average', 'Max', 'Min' (optional, default: 'Sum')",
                @enum = new[] { "Sum", "Count", "Average", "Max", "Min" }
            }
        },
        required = new[] { "path", "pivotTableIndex", "fieldName", "fieldType" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int>() ?? throw new ArgumentException("pivotTableIndex is required");
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required");
        var fieldType = arguments?["fieldType"]?.GetValue<string>() ?? throw new ArgumentException("fieldType is required");
        var function = arguments?["function"]?.GetValue<string>() ?? "Sum";

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var pivotTables = worksheet.PivotTables;

        if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
        {
            throw new ArgumentException($"pivotTableIndex must be between 0 and {pivotTables.Count - 1}");
        }

        var pivotTable = pivotTables[pivotTableIndex];
        
        // Get field index from source data
        var sourceRange = pivotTable.DataSource?.ToString();
        if (string.IsNullOrEmpty(sourceRange))
        {
            throw new ArgumentException("Pivot table data source is not available");
        }
        
        var sourceSheet = workbook.Worksheets[sheetIndex];
        var sourceRangeStr = sourceRange.Replace("=", "");
        var sourceParts = sourceRangeStr.Split(new[] { '!' }, StringSplitOptions.RemoveEmptyEntries);
        var rangeStr = sourceParts.Length > 1 ? sourceParts[1] : sourceParts[0];
        var sourceRangeObj = sourceSheet.Cells.CreateRange(rangeStr);
        
        // Find field index by searching header row
        int fieldIndex = -1;
        for (int col = sourceRangeObj.FirstColumn; col <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount; col++)
        {
            var headerCell = sourceSheet.Cells[sourceRangeObj.FirstRow, col];
            if (headerCell.Value?.ToString() == fieldName)
            {
                fieldIndex = col - sourceRangeObj.FirstColumn;
                break;
            }
        }

        if (fieldIndex < 0)
        {
            throw new ArgumentException($"Field '{fieldName}' not found in pivot table source data");
        }

        switch (fieldType.ToLower())
        {
            case "row":
                pivotTable.AddFieldToArea(PivotFieldType.Row, fieldIndex);
                break;
            case "column":
                pivotTable.AddFieldToArea(PivotFieldType.Column, fieldIndex);
                break;
            case "data":
                pivotTable.AddFieldToArea(PivotFieldType.Data, fieldIndex);
                if (pivotTable.DataFields.Count > 0)
                {
                    var dataField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
                    var functionType = function switch
                    {
                        "Count" => ConsolidationFunction.Count,
                        "Average" => ConsolidationFunction.Average,
                        "Max" => ConsolidationFunction.Max,
                        "Min" => ConsolidationFunction.Min,
                        _ => ConsolidationFunction.Sum
                    };
                    dataField.Function = functionType;
                }
                break;
            case "page":
                pivotTable.AddFieldToArea(PivotFieldType.Page, fieldIndex);
                break;
        }

        pivotTable.RefreshData();
        pivotTable.CalculateData();

        workbook.Save(path);
        return await Task.FromResult($"Field '{fieldName}' added as {fieldType} field to pivot table #{pivotTableIndex}: {path}");
    }
}

