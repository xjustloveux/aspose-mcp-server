using Aspose.Cells;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class for Excel data validation operations.
/// </summary>
public static class ExcelDataValidationHelper
{
    /// <summary>
    ///     Parses validation type string to ValidationType enum.
    /// </summary>
    /// <param name="validationType">The validation type string to parse.</param>
    /// <returns>The corresponding ValidationType enum value.</returns>
    /// <exception cref="ArgumentException">Thrown when the validation type is not supported.</exception>
    public static ValidationType ParseValidationType(string validationType)
    {
        return validationType switch
        {
            "WholeNumber" => ValidationType.WholeNumber,
            "Decimal" => ValidationType.Decimal,
            "List" => ValidationType.List,
            "Date" => ValidationType.Date,
            "Time" => ValidationType.Time,
            "TextLength" => ValidationType.TextLength,
            "Custom" => ValidationType.Custom,
            _ => throw new ArgumentException($"Unsupported validation type: {validationType}")
        };
    }

    /// <summary>
    ///     Parses operator type string to OperatorType enum.
    /// </summary>
    /// <param name="operatorType">The operator type string to parse.</param>
    /// <param name="formula2">The second formula value used to infer operator type if not specified.</param>
    /// <returns>The corresponding OperatorType enum value.</returns>
    /// <exception cref="ArgumentException">Thrown when the operator type is not supported.</exception>
    public static OperatorType ParseOperatorType(string? operatorType, string? formula2)
    {
        if (!string.IsNullOrEmpty(operatorType))
            return operatorType switch
            {
                "Between" => OperatorType.Between,
                "Equal" => OperatorType.Equal,
                "NotEqual" => OperatorType.NotEqual,
                "GreaterThan" => OperatorType.GreaterThan,
                "LessThan" => OperatorType.LessThan,
                "GreaterOrEqual" => OperatorType.GreaterOrEqual,
                "LessOrEqual" => OperatorType.LessOrEqual,
                _ => throw new ArgumentException($"Unsupported operator type: {operatorType}")
            };

        return !string.IsNullOrEmpty(formula2) ? OperatorType.Between : OperatorType.Equal;
    }

    /// <summary>
    ///     Validates collection index and throws exception if invalid.
    /// </summary>
    /// <param name="index">The index to validate.</param>
    /// <param name="count">The total count of items in the collection.</param>
    /// <param name="itemName">The name of the item type for error messages.</param>
    /// <exception cref="ArgumentException">Thrown when the index is out of range.</exception>
    public static void ValidateCollectionIndex(int index, int count, string itemName)
    {
        if (index < 0 || index >= count)
            throw new ArgumentException(
                $"{itemName} index {index} is out of range (collection has {count} {itemName}s)");
    }
}
