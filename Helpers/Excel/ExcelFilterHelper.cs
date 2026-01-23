using Aspose.Cells;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class for Excel filter operations.
/// </summary>
public static class ExcelFilterHelper
{
    /// <summary>
    ///     Determines if the filter operator is a numeric comparison operator.
    /// </summary>
    /// <param name="op">The filter operator type to check.</param>
    /// <returns>True if the operator requires numeric comparison (GreaterThan, LessThan, etc.); otherwise false.</returns>
    public static bool IsNumericOperator(FilterOperatorType op)
    {
        return op is FilterOperatorType.GreaterThan or FilterOperatorType.GreaterOrEqual
            or FilterOperatorType.LessThan or FilterOperatorType.LessOrEqual;
    }

    /// <summary>
    ///     Parses filter operator string to FilterOperatorType enum.
    /// </summary>
    /// <param name="operatorStr">The filter operator string.</param>
    /// <returns>The corresponding FilterOperatorType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the operator string is not supported.</exception>
    public static FilterOperatorType ParseFilterOperator(string operatorStr)
    {
        return operatorStr switch
        {
            "Equal" => FilterOperatorType.Equal,
            "NotEqual" => FilterOperatorType.NotEqual,
            "GreaterThan" => FilterOperatorType.GreaterThan,
            "GreaterOrEqual" => FilterOperatorType.GreaterOrEqual,
            "LessThan" => FilterOperatorType.LessThan,
            "LessOrEqual" => FilterOperatorType.LessOrEqual,
            "Contains" => FilterOperatorType.Contains,
            "NotContains" => FilterOperatorType.NotContains,
            "BeginsWith" => FilterOperatorType.BeginsWith,
            "EndsWith" => FilterOperatorType.EndsWith,
            _ => throw new ArgumentException($"Unsupported filter operator: {operatorStr}")
        };
    }
}
