using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class for Excel conditional formatting operations.
/// </summary>
public static class ExcelConditionalFormattingHelper
{
    /// <summary>
    ///     Regex pattern for validating Excel range format (e.g., A1:B10).
    /// </summary>
    private static readonly Regex RangeRegex = new(@"^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$", RegexOptions.Compiled,
        TimeSpan.FromSeconds(1));

    /// <summary>
    ///     Valid condition types for conditional formatting.
    /// </summary>
    private static readonly string[] ValidConditions = ["greaterthan", "lessthan", "between", "equal"];

    /// <summary>
    ///     Validates the range format (e.g., 'A1:B10').
    /// </summary>
    /// <param name="range">The range string to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the range format is invalid.</exception>
    public static void ValidateRange(string range)
    {
        if (!RangeRegex.IsMatch(range))
            throw new ArgumentException($"Invalid range format: '{range}'. Expected format like 'A1:B10', 'C1:D5'");
    }

    /// <summary>
    ///     Parses condition string to OperatorType.
    /// </summary>
    /// <param name="conditionStr">The condition string to parse (e.g., 'GreaterThan', 'LessThan').</param>
    /// <param name="defaultOperator">The default operator to return if conditionStr is null or empty.</param>
    /// <returns>The corresponding OperatorType enum value.</returns>
    public static OperatorType ParseOperatorType(string? conditionStr,
        OperatorType defaultOperator = OperatorType.GreaterThan)
    {
        if (string.IsNullOrEmpty(conditionStr))
            return defaultOperator;

        return conditionStr.ToLower() switch
        {
            "greaterthan" => OperatorType.GreaterThan,
            "lessthan" => OperatorType.LessThan,
            "between" => OperatorType.Between,
            "equal" => OperatorType.Equal,
            _ => defaultOperator
        };
    }

    /// <summary>
    ///     Checks if the condition string is a valid operator type.
    /// </summary>
    /// <param name="conditionStr">The condition string to check.</param>
    /// <returns>True if the condition string is a valid operator type; otherwise, false.</returns>
    public static bool IsValidCondition(string conditionStr)
    {
        return ValidConditions.Contains(conditionStr.ToLower());
    }
}
