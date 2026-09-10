using Aspose.Words;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Selects the runs that belong to a paragraph itself.
///     <see cref="CompositeNode.GetChildNodes(NodeType,bool)" /> with a deep search also returns the
///     runs inside an inline shape's own paragraphs, so an edit or a delete aimed at the outer
///     paragraph reached into a text box and destroyed content the caller never addressed. A text
///     box paragraph is separately addressable through the Word paragraph tools, so it must not
///     also be part of its host paragraph's character space.
/// </summary>
public static class WordRunHelper
{
    /// <summary>
    ///     Returns the runs whose immediate parent is <paramref name="paragraph" />.
    /// </summary>
    /// <param name="paragraph">The paragraph to read.</param>
    /// <returns>The paragraph's own runs, in document order.</returns>
    public static List<Run> GetDirectRuns(Paragraph paragraph)
    {
        return paragraph.GetChildNodes(NodeType.Run, true)
            .Cast<Run>()
            .Where(run => ReferenceEquals(run.ParentNode, paragraph))
            .ToList();
    }

    /// <summary>
    ///     Total number of characters in a paragraph's own runs. This is the character space the
    ///     range-based Word text operations address.
    /// </summary>
    /// <param name="paragraph">The paragraph to measure.</param>
    /// <returns>The combined length of the paragraph's own runs.</returns>
    public static int GetDirectRunTextLength(Paragraph paragraph)
    {
        return GetDirectRuns(paragraph).Sum(run => run.Text.Length);
    }
}
