using System.Drawing;
using Aspose.Words;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class for Word format operations.
/// </summary>
public static class WordFormatHelper
{
    /// <summary>
    ///     Gets target paragraph using flat list.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index, or -1 for the last paragraph.</param>
    /// <returns>The target paragraph.</returns>
    /// <exception cref="ArgumentException">Thrown when the document has no paragraphs or the index is out of range.</exception>
    public static WordParagraph GetTargetParagraph(Document doc, int paragraphIndex)
    {
        var allParas = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        if (allParas.Count == 0)
            throw new ArgumentException("Document has no paragraphs.");

        if (paragraphIndex == -1)
            return allParas[^1]; // Last paragraph

        if (paragraphIndex < 0 || paragraphIndex >= allParas.Count)
            throw new ArgumentException(
                $"paragraphIndex must be between 0 and {allParas.Count - 1}, or use -1 for last paragraph");

        return allParas[paragraphIndex];
    }

    /// <summary>
    ///     Converts line style string to LineStyle enum.
    /// </summary>
    /// <param name="style">The line style string (none, single, double, dotted, dashed, thick).</param>
    /// <returns>The corresponding LineStyle enum value.</returns>
    public static LineStyle GetLineStyle(string style)
    {
        return style.ToLower() switch
        {
            "none" => LineStyle.None,
            "single" => LineStyle.Single,
            "double" => LineStyle.Double,
            "dotted" => LineStyle.Dot,
            "dashed" => LineStyle.Single,
            "thick" => LineStyle.Thick,
            _ => LineStyle.Single
        };
    }

    /// <summary>
    ///     Gets human-readable color name.
    /// </summary>
    /// <param name="color">The color to get the name for.</param>
    /// <returns>The human-readable color name.</returns>
    public static string GetColorName(Color color)
    {
        if (color is { IsEmpty: true } or { R: 0, G: 0, B: 0, A: 0 })
            return "Auto/Black";

        if (color is { R: 255, G: 0, B: 0 }) return "Red";
        if (color is { R: 0, G: 255, B: 0 }) return "Green";
        if (color is { R: 0, G: 0, B: 255 }) return "Blue";
        if (color is { R: 255, G: 255, B: 0 }) return "Yellow";
        if (color is { R: 255, G: 0, B: 255 }) return "Magenta";
        if (color is { R: 0, G: 255, B: 255 }) return "Cyan";
        if (color is { R: 255, G: 255, B: 255 }) return "White";
        if (color is { R: 128, G: 128, B: 128 }) return "Gray";
        if (color is { R: 255, G: 165, B: 0 }) return "Orange";
        if (color is { R: 128, G: 0, B: 128 }) return "Purple";

        // Try to get the named color
        if (color.IsKnownColor)
            return color.Name;

        return "Custom";
    }
}
