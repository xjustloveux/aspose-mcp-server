using System.Text.RegularExpressions;

namespace AsposeMcpServer.Handlers.Excel.Comment;

/// <summary>
///     Helper class for Excel comment operations.
/// </summary>
public static class ExcelCommentHelper
{
    /// <summary>
    ///     Default author name for comments.
    /// </summary>
    public const string DefaultAuthor = "AsposeMCP";

    /// <summary>
    ///     Regex pattern for validating Excel cell addresses (e.g., A1, B2, AA100).
    /// </summary>
    private static readonly Regex CellAddressRegex =
        new(@"^[A-Za-z]{1,3}\d+$", RegexOptions.Compiled, TimeSpan.FromSeconds(1));

    /// <summary>
    ///     Validates the cell address format.
    /// </summary>
    /// <param name="cell">The cell address to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the cell address format is invalid.</exception>
    public static void ValidateCellAddress(string cell)
    {
        if (!CellAddressRegex.IsMatch(cell))
            throw new ArgumentException(
                $"Invalid cell address format: '{cell}'. Expected format like 'A1', 'B2', 'AA100'");
    }
}
