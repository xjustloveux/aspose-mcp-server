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
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(5);

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
        var searchParams = ExtractSearchParameters(parameters);
        var doc = context.Document;
        var matches = FindAllMatches(doc, searchParams);
        return BuildSearchResults(matches, searchParams);
    }

    /// <summary>
    ///     Extracts search parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted search parameters.</returns>
    private static SearchParameters ExtractSearchParameters(OperationParameters parameters)
    {
        return new SearchParameters(
            parameters.GetRequired<string>("searchText"),
            parameters.GetOptional("useRegex", false),
            parameters.GetOptional("caseSensitive", false),
            parameters.GetOptional("maxResults", 50),
            parameters.GetOptional("contextLength", 50)
        );
    }

    /// <summary>
    ///     Finds all text matches in the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="p">The search parameters.</param>
    /// <returns>A list of matches with text, paragraph index, and context.</returns>
    private static List<(string text, int paragraphIndex, string context)> FindAllMatches(Document doc,
        SearchParameters p)
    {
        List<(string text, int paragraphIndex, string context)> matches = [];
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        for (var i = 0; i < paragraphs.Count && matches.Count < p.MaxResults; i++)
        {
            if (paragraphs[i] is not WordParagraph para) continue;
            var paraText = para.GetText();

            if (p.UseRegex)
                FindRegexMatches(paraText, i, p, matches);
            else
                FindLiteralMatches(paraText, i, p, matches);
        }

        return matches;
    }

    /// <summary>
    ///     Finds matches using regular expression.
    /// </summary>
    /// <param name="paraText">The paragraph text.</param>
    /// <param name="paraIndex">The paragraph index.</param>
    /// <param name="p">The search parameters.</param>
    /// <param name="matches">The list to add matches to.</param>
    private static void FindRegexMatches(string paraText, int paraIndex, SearchParameters p,
        List<(string text, int paragraphIndex, string context)> matches)
    {
        var options = p.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
        var regex = new Regex(p.SearchText, options, RegexTimeout);

        foreach (Match match in regex.Matches(paraText))
        {
            if (matches.Count >= p.MaxResults) break;
            var ctx = GetContext(paraText, match.Index, match.Length, p.ContextLength);
            matches.Add((match.Value, paraIndex, ctx));
        }
    }

    /// <summary>
    ///     Finds matches using literal string comparison.
    /// </summary>
    /// <param name="paraText">The paragraph text.</param>
    /// <param name="paraIndex">The paragraph index.</param>
    /// <param name="p">The search parameters.</param>
    /// <param name="matches">The list to add matches to.</param>
    private static void FindLiteralMatches(string paraText, int paraIndex, SearchParameters p,
        List<(string text, int paragraphIndex, string context)> matches)
    {
        var comparison = p.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = 0;

        while ((index = paraText.IndexOf(p.SearchText, index, comparison)) != -1)
        {
            if (matches.Count >= p.MaxResults) break;
            var ctx = GetContext(paraText, index, p.SearchText.Length, p.ContextLength);
            matches.Add((p.SearchText, paraIndex, ctx));
            index += p.SearchText.Length;
        }
    }

    /// <summary>
    ///     Builds the formatted search results.
    /// </summary>
    /// <param name="matches">The list of matches.</param>
    /// <param name="p">The search parameters.</param>
    /// <returns>The formatted search results.</returns>
    private static string BuildSearchResults(List<(string text, int paragraphIndex, string context)> matches,
        SearchParameters p)
    {
        var result = new StringBuilder();
        result.AppendLine("=== Search Results ===");
        result.AppendLine($"Search text: {p.SearchText}");
        result.AppendLine($"Use regex: {(p.UseRegex ? "Yes" : "No")}");
        result.AppendLine($"Case sensitive: {(p.CaseSensitive ? "Yes" : "No")}");
        result.AppendLine(
            $"Found {matches.Count} matches{(matches.Count >= p.MaxResults ? $" (limited to first {p.MaxResults})" : "")}\n");

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

    /// <summary>
    ///     Record to hold search parameters.
    /// </summary>
    private sealed record SearchParameters(
        string SearchText,
        bool UseRegex,
        bool CaseSensitive,
        int MaxResults,
        int ContextLength);
}
