using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     A resolved paragraph: the live node, its normalized address, and its global document-order
///     index (position among all paragraphs in document order, across every story).
/// </summary>
/// <param name="Paragraph">The resolved paragraph node.</param>
/// <param name="Address">The normalized address (e.g. -1 resolved to the concrete last index).</param>
/// <param name="DocumentOrderIndex">The paragraph's index in document order across all stories.</param>
public sealed record ParagraphRef(WordParagraph Paragraph, ParagraphAddress Address, int DocumentOrderIndex);
