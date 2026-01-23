using System.Text;
using Aspose.Slides;

namespace AsposeMcpServer.Helpers.PowerPoint;

/// <summary>
///     Helper class providing shared text manipulation methods for PowerPoint text handlers.
/// </summary>
public static class PptTextHelper
{
    /// <summary>
    ///     Recursively processes shapes for text replacement, including GroupShapes and Tables.
    /// </summary>
    /// <param name="shapes">The shape collection to process.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="comparison">The string comparison type.</param>
    /// <returns>The number of replacements made.</returns>
    public static int ProcessShapesForReplace(IShapeCollection shapes, string findText, string replaceText,
        StringComparison comparison)
    {
        var replacements = 0;

        foreach (var shape in shapes)
            switch (shape)
            {
                case IAutoShape { TextFrame: not null } autoShape:
                    replacements += ReplaceInTextFrame(autoShape.TextFrame, findText, replaceText, comparison);
                    break;
                case IGroupShape groupShape:
                    replacements += ProcessShapesForReplace(groupShape.Shapes, findText, replaceText, comparison);
                    break;
                case ITable table:
                    replacements += ReplaceInTable(table, findText, replaceText, comparison);
                    break;
            }

        return replacements;
    }

    /// <summary>
    ///     Replaces text in a TextFrame while preserving formatting at the Portion level.
    /// </summary>
    /// <param name="textFrame">The text frame to process.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="comparison">The string comparison type.</param>
    /// <returns>The number of replacements made (0 or 1).</returns>
    private static int ReplaceInTextFrame(ITextFrame textFrame, string findText, string replaceText,
        StringComparison comparison)
    {
        var originalText = textFrame.Text;
        if (string.IsNullOrEmpty(originalText)) return 0;

        if (originalText.IndexOf(findText, comparison) < 0) return 0;

        foreach (var para in textFrame.Paragraphs)
        foreach (var portion in para.Portions)
        {
            var portionText = portion.Text;
            if (string.IsNullOrEmpty(portionText)) continue;

            var newText = ReplaceAll(portionText, findText, replaceText, comparison);
            if (newText != portionText)
                portion.Text = newText;
        }

        return 1;
    }

    /// <summary>
    ///     Replaces text in all cells of a table.
    /// </summary>
    /// <param name="table">The table to process.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="comparison">The string comparison type.</param>
    /// <returns>The number of replacements made.</returns>
    private static int ReplaceInTable(ITable table, string findText, string replaceText, StringComparison comparison)
    {
        var replacements = 0;

        for (var row = 0; row < table.Rows.Count; row++)
        for (var col = 0; col < table.Columns.Count; col++)
        {
            var cell = table[col, row];
            if (cell.TextFrame != null)
                replacements += ReplaceInTextFrame(cell.TextFrame, findText, replaceText, comparison);
        }

        return replacements;
    }

    /// <summary>
    ///     Replaces all occurrences of a string in source with replacement string.
    /// </summary>
    /// <param name="source">The source string to search in.</param>
    /// <param name="find">The text to find.</param>
    /// <param name="replace">The text to replace with.</param>
    /// <param name="comparison">The string comparison type.</param>
    /// <returns>The string with all occurrences replaced.</returns>
    private static string ReplaceAll(string source, string find, string replace, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(find)) return source;

        var sb = new StringBuilder();
        var idx = 0;
        while (true)
        {
            var next = source.IndexOf(find, idx, comparison);
            if (next < 0)
            {
                sb.Append(source, idx, source.Length - idx);
                break;
            }

            sb.Append(source, idx, next - idx);
            sb.Append(replace);
            idx = next + find.Length;
        }

        return sb.ToString();
    }
}
