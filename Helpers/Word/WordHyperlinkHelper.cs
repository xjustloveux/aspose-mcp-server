using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper methods for Word hyperlink operations.
/// </summary>
public static class WordHyperlinkHelper
{
    /// <summary>
    ///     Gets all hyperlink fields from the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <returns>A list of FieldHyperlink objects representing all hyperlinks.</returns>
    public static List<FieldHyperlink> GetAllHyperlinks(Document doc)
    {
        return doc.Range.Fields.OfType<FieldHyperlink>().ToList();
    }

    /// <summary>
    ///     Validates URL format to prevent invalid field commands.
    /// </summary>
    /// <param name="url">The URL to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the URL format is invalid.</exception>
    public static void ValidateUrlFormat(string url)
    {
        var validPrefixes = new[] { "http://", "https://", "mailto:", "ftp://", "file://", "#" };
        if (!validPrefixes.Any(prefix => url.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException(
                $"Invalid URL format: '{url}'. URL must start with http://, https://, mailto:, ftp://, file://, or # (for internal links)");
    }
}
