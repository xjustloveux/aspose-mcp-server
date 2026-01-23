using Aspose.Words;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class for Word paragraph operations.
/// </summary>
public static class WordParagraphHelper
{
    /// <summary>
    ///     Converts an alignment string to ParagraphAlignment enum.
    /// </summary>
    /// <param name="alignment">Alignment string (left, center, right, justify).</param>
    /// <returns>Corresponding ParagraphAlignment value.</returns>
    public static ParagraphAlignment GetAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => ParagraphAlignment.Left,
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            "justify" => ParagraphAlignment.Justify,
            _ => ParagraphAlignment.Left
        };
    }

    /// <summary>
    ///     Converts a line spacing rule string to LineSpacingRule enum.
    /// </summary>
    /// <param name="rule">Line spacing rule string (atleast, exactly, or default multiple).</param>
    /// <returns>Corresponding LineSpacingRule value.</returns>
    public static LineSpacingRule GetLineSpacingRule(string rule)
    {
        return rule.ToLower() switch
        {
            "atleast" => LineSpacingRule.AtLeast,
            "exactly" => LineSpacingRule.Exactly,
            _ => LineSpacingRule.Multiple
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
            _ => TabAlignment.Left
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
            _ => TabLeader.None
        };
    }
}
