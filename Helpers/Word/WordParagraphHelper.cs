using Aspose.Words;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class for Word paragraph operations.
/// </summary>
public static class WordParagraphHelper
{
    /// <summary>
    ///     Converts a line spacing rule string to LineSpacingRule enum.
    /// </summary>
    /// <summary>The line spacing rules <c>word_paragraph</c> documents, in its own wording.</summary>
    private const string LineSpacingRules = "single, oneAndHalf, double, atLeast, exactly, multiple";

    /// <summary>
    ///     Converts an alignment string to ParagraphAlignment enum.
    /// </summary>
    /// <param name="alignment">Alignment string (left, center, right, justify).</param>
    /// <returns>Corresponding ParagraphAlignment value.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the value is not one this helper knows. It used to return Left, so a typo
    ///     became a silent left-alignment and the request reported success (§23.13.1).
    /// </exception>
    public static ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" => ParagraphAlignment.Justify,
            _ => throw new ArgumentException(
                $"Unknown alignment '{alignment}'. Use left, center, right or justify.")
        };
    }

    /// <param name="rule">
    ///     Line spacing rule string. The multiple-based rules — single, oneAndHalf, double and
    ///     multiple — all map to <see cref="LineSpacingRule.Multiple" />; which multiple is
    ///     decided by the accompanying spacing value.
    /// </param>
    /// <returns>Corresponding LineSpacingRule value.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the value is not one the tool documents. It used to return Multiple, so a
    ///     typo became silent single spacing (§23.13.1).
    /// </exception>
    public static LineSpacingRule GetLineSpacingRule(string rule)
    {
        return rule.ToLower() switch
        {
            "atleast" => LineSpacingRule.AtLeast,
            "exactly" => LineSpacingRule.Exactly,
            "single" or "oneandhalf" or "double" or "multiple" => LineSpacingRule.Multiple,
            _ => throw new ArgumentException(
                $"Unknown lineSpacingRule '{rule}'. Use one of: {LineSpacingRules}.")
        };
    }

    /// <summary>
    ///     Converts a tab alignment string to TabAlignment enum.
    /// </summary>
    /// <param name="alignment">Tab alignment string (left, center, right, decimal, bar, clear).</param>
    /// <returns>Corresponding TabAlignment value.</returns>
    public static TabAlignment GetTabAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => TabAlignment.Left,
            "center" => TabAlignment.Center,
            "right" => TabAlignment.Right,
            "decimal" => TabAlignment.Decimal,
            "bar" => TabAlignment.Bar,
            "clear" => TabAlignment.Clear,
            _ => throw new ArgumentException(
                $"Unknown tab alignment '{alignment}'. Use left, center, right, decimal, bar or "
                + "clear.")
        };
    }

    /// <summary>
    ///     Converts a tab leader string to TabLeader enum.
    /// </summary>
    /// <param name="leader">Tab leader string (none, dots, dashes, line, heavy, middledot).</param>
    /// <returns>Corresponding TabLeader value.</returns>
    public static TabLeader GetTabLeader(string leader)
    {
        return leader.ToLower() switch
        {
            "none" => TabLeader.None,
            "dots" => TabLeader.Dots,
            "dashes" => TabLeader.Dashes,
            "line" => TabLeader.Line,
            "heavy" => TabLeader.Heavy,
            "middledot" => TabLeader.MiddleDot,
            _ => throw new ArgumentException(
                $"Unknown tab leader '{leader}'. Use none, dots, dashes, line, heavy or middledot.")
        };
    }
}
