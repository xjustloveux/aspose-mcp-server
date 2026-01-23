using Aspose.Cells;
using Aspose.Cells.Pivot;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class for Excel pivot table operations.
/// </summary>
public static class ExcelPivotTableHelper
{
    /// <summary>
    ///     Parses a style name string to PivotTableStyleType enum.
    /// </summary>
    /// <param name="style">The style name string to parse.</param>
    /// <returns>The corresponding PivotTableStyleType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the style name is invalid.</exception>
    public static PivotTableStyleType ParsePivotTableStyle(string style)
    {
        if (string.Equals(style, "None", StringComparison.OrdinalIgnoreCase))
            return PivotTableStyleType.None;

        if (Enum.TryParse<PivotTableStyleType>($"PivotTableStyle{style}", true, out var result))
            return result;

        if (Enum.TryParse(style, true, out result))
            return result;

        throw new ArgumentException(
            $"Invalid style: '{style}'. Valid formats: 'Light1'-'Light28', 'Medium1'-'Medium28', 'Dark1'-'Dark28', or 'None'");
    }

    /// <summary>
    ///     Finds field index in pivot table source data.
    /// </summary>
    /// <param name="sourceSheet">The source worksheet.</param>
    /// <param name="sourceRangeObj">The source range object.</param>
    /// <param name="fieldName">The field name to find.</param>
    /// <returns>The field index or -1 if not found.</returns>
    public static int FindFieldIndex(Worksheet sourceSheet, Range sourceRangeObj, string fieldName)
    {
        var fieldIndex = FindFieldInHeaderRow(sourceSheet, sourceRangeObj, fieldName);
        return fieldIndex >= 0 ? fieldIndex : FindFieldInAllCells(sourceSheet, sourceRangeObj, fieldName);
    }

    /// <summary>
    ///     Finds field index by searching the header row only.
    /// </summary>
    /// <param name="sourceSheet">The source worksheet.</param>
    /// <param name="sourceRangeObj">The source range object.</param>
    /// <param name="fieldName">The field name to find.</param>
    /// <returns>The field index or -1 if not found.</returns>
    private static int FindFieldInHeaderRow(Worksheet sourceSheet, Range sourceRangeObj, string fieldName)
    {
        var headerRowIndex = sourceRangeObj.FirstRow;
        var trimmedFieldName = fieldName.Trim();

        for (var col = sourceRangeObj.FirstColumn;
             col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
             col++)
        {
            var cellValue = sourceSheet.Cells[headerRowIndex, col].Value?.ToString()?.Trim();
            if (cellValue == fieldName || cellValue == trimmedFieldName)
                return col - sourceRangeObj.FirstColumn;
        }

        return -1;
    }

    /// <summary>
    ///     Finds field index by searching all cells in the source range.
    /// </summary>
    /// <param name="sourceSheet">The source worksheet.</param>
    /// <param name="sourceRangeObj">The source range object.</param>
    /// <param name="fieldName">The field name to find.</param>
    /// <returns>The field index or -1 if not found.</returns>
    private static int FindFieldInAllCells(Worksheet sourceSheet, Range sourceRangeObj, string fieldName)
    {
        var trimmedFieldName = fieldName.Trim();

        for (var row = sourceRangeObj.FirstRow;
             row < sourceRangeObj.FirstRow + sourceRangeObj.RowCount;
             row++)
        for (var col = sourceRangeObj.FirstColumn;
             col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
             col++)
        {
            var cellValue = sourceSheet.Cells[row, col].Value?.ToString()?.Trim();
            if (cellValue == fieldName || cellValue == trimmedFieldName)
                return col - sourceRangeObj.FirstColumn;
        }

        return -1;
    }

    /// <summary>
    ///     Gets available field names from source data.
    /// </summary>
    /// <param name="sourceSheet">The source worksheet.</param>
    /// <param name="sourceRangeObj">The source range object.</param>
    /// <returns>List of available field names.</returns>
    public static List<string> GetAvailableFieldNames(Worksheet sourceSheet, Range sourceRangeObj)
    {
        List<string> availableFields = [];
        var headerRowIndex = sourceRangeObj.FirstRow;

        for (var col = sourceRangeObj.FirstColumn;
             col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
             col++)
        {
            var headerCell = sourceSheet.Cells[headerRowIndex, col];
            var cellValue = headerCell.Value?.ToString()?.Trim();
            if (!string.IsNullOrEmpty(cellValue))
                availableFields.Add(cellValue);
        }

        return availableFields;
    }

    /// <summary>
    ///     Parses the data source to get range string.
    /// </summary>
    /// <param name="workbook">The workbook containing the pivot table.</param>
    /// <param name="pivotTable">The pivot table.</param>
    /// <param name="sheetIndex">The sheet index for resolving the source range.</param>
    /// <param name="pivotTableIndex">The pivot table index for error messages.</param>
    /// <param name="worksheetName">The worksheet name for error messages.</param>
    /// <returns>The clean source range string.</returns>
    public static (Worksheet sourceSheet, Range sourceRangeObj) ParseDataSource(
        Workbook workbook,
        PivotTable pivotTable,
        int sheetIndex,
        int pivotTableIndex,
        string worksheetName)
    {
        string? sourceRangeStr = null;
        var dataSource = pivotTable.DataSource;

        if (dataSource is Array { Length: > 0 } dataSourceArray)
            sourceRangeStr = dataSourceArray.GetValue(0)?.ToString();
        else if (dataSource != null) sourceRangeStr = dataSource.ToString();

        if (string.IsNullOrEmpty(sourceRangeStr)) sourceRangeStr = pivotTable.DataSource?.ToString();

        if (string.IsNullOrEmpty(sourceRangeStr))
            throw new ArgumentException(
                $"Pivot table data source is not available. Pivot table index: {pivotTableIndex}, Worksheet: '{worksheetName}'");

        var sourceSheet = workbook.Worksheets[sheetIndex];
        var cleanSourceRange = sourceRangeStr.Replace("=", "").Trim();
        var sourceParts = cleanSourceRange.Split(['!'], StringSplitOptions.RemoveEmptyEntries);
        var rangeStr = sourceParts.Length > 1 ? sourceParts[1].Trim() : sourceParts[0].Trim();

        if (string.IsNullOrEmpty(rangeStr))
            throw new ArgumentException(
                $"Invalid data source format: '{sourceRangeStr}'. Unable to parse range from data source. Pivot table index: {pivotTableIndex}, Worksheet: '{worksheetName}'");

        Range sourceRangeObj;
        try
        {
            sourceRangeObj = sourceSheet.Cells.CreateRange(rangeStr);
        }
        catch (Exception rangeEx)
        {
            throw new ArgumentException(
                $"Failed to parse pivot table data source range '{rangeStr}' from source '{sourceRangeStr}': {rangeEx.Message}");
        }

        return (sourceSheet, sourceRangeObj);
    }

    /// <summary>
    ///     Parses field type string to PivotFieldType enum.
    /// </summary>
    public static PivotFieldType ParseFieldType(string fieldType)
    {
        return fieldType.ToLower() switch
        {
            "row" => PivotFieldType.Row,
            "column" => PivotFieldType.Column,
            "data" => PivotFieldType.Data,
            "page" => PivotFieldType.Page,
            _ => throw new ArgumentException(
                $"Invalid fieldType: {fieldType}. Valid values are: Row, Column, Data, Page")
        };
    }

    /// <summary>
    ///     Parses function string to ConsolidationFunction enum.
    /// </summary>
    public static ConsolidationFunction ParseFunction(string function)
    {
        return function switch
        {
            "Count" => ConsolidationFunction.Count,
            "Average" => ConsolidationFunction.Average,
            "Max" => ConsolidationFunction.Max,
            "Min" => ConsolidationFunction.Min,
            _ => ConsolidationFunction.Sum
        };
    }
}
