using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class for Excel group operations.
/// </summary>
public static class ExcelGroupHelper
{
    /// <summary>
    ///     Validates that required parameters are present for an operation.
    /// </summary>
    /// <param name="operation">The operation name.</param>
    /// <param name="parameters">The parameters dictionary.</param>
    /// <param name="requiredParams">The required parameter names.</param>
    /// <exception cref="ArgumentException">Thrown when a required parameter is missing.</exception>
    public static void ValidateRequiredParams(string operation, OperationParameters parameters,
        params string[] requiredParams)
    {
        var missingParam = requiredParams.FirstOrDefault(param => !parameters.Has(param));
        if (missingParam != null)
            throw new ArgumentException($"Operation '{operation}' requires parameter '{missingParam}'.");
    }

    /// <summary>
    ///     Validates row range indices.
    /// </summary>
    /// <param name="startRow">The start row index (0-based).</param>
    /// <param name="endRow">The end row index (0-based).</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when startRow or endRow is negative, or when startRow is greater than
    ///     endRow.
    /// </exception>
    public static void ValidateRowRange(int startRow, int endRow)
    {
        if (startRow < 0)
            throw new ArgumentException($"startRow cannot be negative. Got: {startRow}");
        if (endRow < 0)
            throw new ArgumentException($"endRow cannot be negative. Got: {endRow}");
        if (startRow > endRow)
            throw new ArgumentException($"startRow ({startRow}) cannot be greater than endRow ({endRow}).");
    }

    /// <summary>
    ///     Validates column range indices.
    /// </summary>
    /// <param name="startColumn">The start column index (0-based).</param>
    /// <param name="endColumn">The end column index (0-based).</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when startColumn or endColumn is negative, or when startColumn is greater
    ///     than endColumn.
    /// </exception>
    public static void ValidateColumnRange(int startColumn, int endColumn)
    {
        if (startColumn < 0)
            throw new ArgumentException($"startColumn cannot be negative. Got: {startColumn}");
        if (endColumn < 0)
            throw new ArgumentException($"endColumn cannot be negative. Got: {endColumn}");
        if (startColumn > endColumn)
            throw new ArgumentException($"startColumn ({startColumn}) cannot be greater than endColumn ({endColumn}).");
    }
}
