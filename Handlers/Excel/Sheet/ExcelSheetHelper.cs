using System.Buffers;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Helper class for Excel sheet operations providing validation and utility methods.
/// </summary>
public static class ExcelSheetHelper
{
    /// <summary>
    ///     Characters that are not allowed in Excel sheet names.
    /// </summary>
    private static readonly SearchValues<char> InvalidSheetNameChars =
        SearchValues.Create(['\\', '/', '?', '*', '[', ']', ':']);

    /// <summary>
    ///     Validates that the sheet name meets Excel requirements.
    /// </summary>
    /// <param name="name">The sheet name to validate.</param>
    /// <param name="paramName">The parameter name for error messages.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the sheet name is empty, exceeds 31 characters, or contains invalid characters.
    /// </exception>
    public static void ValidateSheetName(string name, string paramName)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException($"{paramName} cannot be empty");

        if (name.Length > 31)
            throw new ArgumentException(
                $"{paramName} '{name}' (length: {name.Length}) exceeds Excel's limit of 31 characters");

        var invalidCharIndex = name.AsSpan().IndexOfAny(InvalidSheetNameChars);
        if (invalidCharIndex >= 0)
            throw new ArgumentException(
                $"{paramName} contains invalid character '{name[invalidCharIndex]}'. Sheet names cannot contain: \\ / ? * [ ] :");
    }
}
