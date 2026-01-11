using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for searching text in Word documents.
/// </summary>
public class SearchWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "search";

    /// <summary>
    ///     Searches for text in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: searchText.
    ///     Optional: useRegex, caseSensitive, maxResults, contextLength.
    /// </param>
    /// <returns>Search results as formatted text.</returns>
    /// <exception cref="ArgumentException">Thrown when searchText is missing.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var searchText = parameters.GetRequired<string>("searchText");
        var useRegex = parameters.GetOptional("useRegex", false);
        var caseSensitive = parameters.GetOptional("caseSensitive", false);
        var maxResults = parameters.GetOptional("maxResults", 50);
        var contextLength = parameters.GetOptional("contextLength", 50);

        var doc = context.Document;
        var result = new StringBuilder();
        List<(string text, int paragraphIndex, string context)> matches = [];

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        for (var i = 0; i < paragraphs.Count && matches.Count < maxResults; i++)
        {
            if (paragraphs[i] is not WordParagraph para) continue;

            var paraText = para.GetText();

            if (useRegex)
            {
                var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
                var regex = new Regex(searchText, options);
                var regexMatches = regex.Matches(paraText);

                foreach (Match match in regexMatches)
                {
                    if (matches.Count >= maxResults) break;

                    var ctx = GetContext(paraText, match.Index, match.Length, contextLength);
                    matches.Add((match.Value, i, ctx));
                }
            }
            else
            {
                var comparison = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                var index = 0;

                while ((index = paraText.IndexOf(searchText, index, comparison)) != -1)
                {
                    if (matches.Count >= maxResults) break;

                    var ctx = GetContext(paraText, index, searchText.Length, contextLength);
                    matches.Add((searchText, i, ctx));
                    index += searchText.Length;
                }
            }
        }

        result.AppendLine("=== Search Results ===");
        result.AppendLine($"Search text: {searchText}");
        result.AppendLine($"Use regex: {(useRegex ? "Yes" : "No")}");
        result.AppendLine($"Case sensitive: {(caseSensitive ? "Yes" : "No")}");
        result.AppendLine(
            $"Found {matches.Count} matches{(matches.Count >= maxResults ? $" (limited to first {maxResults})" : "")}\n");

        if (matches.Count == 0)
            result.AppendLine("No matching text found");
        else
            for (var i = 0; i < matches.Count; i++)
            {
                var match = matches[i];
                result.AppendLine($"Match #{i + 1}:");
                result.AppendLine($"  Location: Paragraph #{match.paragraphIndex}");
                result.AppendLine($"  Matched text: {match.text}");
                result.AppendLine($"  Context: ...{match.context}...");
                result.AppendLine();
            }

        return result.ToString();
    }

    /// <summary>
    ///     Extracts context around a matched text.
    /// </summary>
    /// <param name="text">The full paragraph text.</param>
    /// <param name="matchIndex">Starting index of the match.</param>
    /// <param name="matchLength">Length of the matched text.</param>
    /// <param name="contextLength">Number of characters to include before and after.</param>
    /// <returns>Context string with the match highlighted using brackets.</returns>
    private static string GetContext(string text, int matchIndex, int matchLength, int contextLength)
    {
        var start = Math.Max(0, matchIndex - contextLength);
        var end = Math.Min(text.Length, matchIndex + matchLength + contextLength);

        var context = text.Substring(start, end - start);
        context = context.Replace("\r", "").Replace("\n", " ").Trim();

        var highlightStart = matchIndex - start;
        var highlightEnd = highlightStart + matchLength;

        if (highlightStart >= 0 && highlightEnd <= context.Length)
            context = context.Substring(0, highlightStart) +
                      "[" + context.Substring(highlightStart, matchLength) + "]" +
                      context.Substring(highlightEnd);

        return context;
    }
}
