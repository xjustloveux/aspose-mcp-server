using Aspose.Words;
using Aspose.Words.Fields;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class providing shared methods for Word header/footer handlers.
/// </summary>
public static class WordHeaderFooterHelper
{
    private static readonly Dictionary<string, FieldType> FieldCodeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["PAGE"] = FieldType.FieldPage,
        ["NUMPAGES"] = FieldType.FieldNumPages,
        ["DATE"] = FieldType.FieldDate,
        ["TIME"] = FieldType.FieldTime,
        ["FILENAME"] = FieldType.FieldFileName,
        ["AUTHOR"] = FieldType.FieldAuthor,
        ["TITLE"] = FieldType.FieldTitle
    };

    /// <summary>
    ///     Gets the HeaderFooterType based on the type string.
    /// </summary>
    /// <param name="headerFooterType">The type string (primary, first, even).</param>
    /// <param name="isHeader">True for header, false for footer.</param>
    /// <returns>The corresponding HeaderFooterType enum value.</returns>
    public static HeaderFooterType GetHeaderFooterType(string headerFooterType, bool isHeader)
    {
        return headerFooterType.ToLower() switch
        {
            "first" => isHeader ? HeaderFooterType.HeaderFirst : HeaderFooterType.FooterFirst,
            "even" => isHeader ? HeaderFooterType.HeaderEven : HeaderFooterType.FooterEven,
            _ => isHeader ? HeaderFooterType.HeaderPrimary : HeaderFooterType.FooterPrimary
        };
    }

    /// <summary>
    ///     Gets an existing header/footer or creates a new one if it doesn't exist.
    /// </summary>
    /// <param name="section">The section to get the header/footer from.</param>
    /// <param name="doc">The document.</param>
    /// <param name="hfType">The header/footer type.</param>
    /// <returns>The header/footer node.</returns>
    public static HeaderFooter GetOrCreateHeaderFooter(Section section, Document doc,
        HeaderFooterType hfType)
    {
        var headerFooter = section.HeadersFooters[hfType];
        if (headerFooter == null)
        {
            headerFooter = new HeaderFooter(doc, hfType);
            section.HeadersFooters.Add(headerFooter);
        }

        return headerFooter;
    }

    /// <summary>
    ///     Clears a header or footer according to the caller's request.
    ///     The three set-text handlers each carried a copy of this logic in which the unconditional
    ///     <c>RemoveAllChildren</c> sat outside the <c>clearExisting</c> test, so the default
    ///     <c>clearTextOnly = false</c> wiped the header even when the caller had asked to keep it.
    /// </summary>
    /// <param name="headerFooter">The header or footer to clear.</param>
    /// <param name="clearExisting">Whether the caller asked for existing content to be removed.</param>
    /// <param name="clearTextOnly">
    ///     When <c>true</c>, run text is emptied but structure, fields and shapes stay in place;
    ///     when <c>false</c>, the whole content is removed. Only consulted when
    ///     <paramref name="clearExisting" /> is <c>true</c>.
    /// </param>
    public static void Clear(HeaderFooter headerFooter, bool clearExisting, bool clearTextOnly)
    {
        if (!clearExisting) return;

        if (clearTextOnly)
            ClearTextOnly(headerFooter);
        else
            headerFooter.RemoveAllChildren();
    }

    /// <summary>
    ///     Clears only the text content from a header/footer, preserving other elements.
    /// </summary>
    /// <param name="headerFooter">The header/footer to clear text from.</param>
    public static void ClearTextOnly(HeaderFooter headerFooter)
    {
        var paragraphs = headerFooter.GetChildNodes(NodeType.Paragraph, true);
        foreach (var para in paragraphs.OfType<WordParagraph>())
        {
            var runs = para.GetChildNodes(NodeType.Run, true);
            foreach (var run in runs.OfType<Run>()) run.Text = string.Empty;
        }
    }

    /// <summary>
    ///     Inserts text or a field code into the document.
    ///     Supports mixed content like "Page {PAGE} of {NUMPAGES}".
    ///     <para>
    ///         Only the documented fields are accepted: PAGE, NUMPAGES, DATE, TIME, FILENAME,
    ///         AUTHOR and TITLE. Any other code is refused rather than inserted verbatim, because
    ///         a live field can reach the filesystem or the network on the next UpdateFields().
    ///     </para>
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="text">
    ///     The text to insert. Field codes are enclosed in braces (e.g., {PAGE}, {DATE}).
    ///     Use {{ and }} for literal braces.
    /// </param>
    /// <param name="fontSettings">The font settings to apply.</param>
    public static void InsertTextOrField(DocumentBuilder builder, string text, FontSettings fontSettings)
    {
        ApplyFontSettings(builder, fontSettings);

        if (!text.Contains('{') && !text.Contains('}'))
        {
            builder.Write(text);
            return;
        }

        InsertMixedContent(builder, text);
    }

    /// <summary>
    ///     Parses and inserts mixed content containing text and field codes.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="text">The text containing potential field codes.</param>
    private static void InsertMixedContent(DocumentBuilder builder, string text)
    {
        var i = 0;
        while (i < text.Length)
            if (text[i] == '{')
            {
                if (i + 1 < text.Length && text[i + 1] == '{')
                {
                    builder.Write("{");
                    i += 2;
                    continue;
                }

                var closingIndex = text.IndexOf('}', i + 1);
                if (closingIndex == -1)
                {
                    builder.Write(text[i..]);
                    break;
                }

                var fieldCode = text[(i + 1)..closingIndex].Trim();
                if (!string.IsNullOrEmpty(fieldCode))
                    InsertFieldByCode(builder, fieldCode);

                i = closingIndex + 1;
            }
            else if (text[i] == '}')
            {
                if (i + 1 < text.Length && text[i + 1] == '}')
                {
                    builder.Write("}");
                    i += 2;
                    continue;
                }

                builder.Write("}");
                i++;
            }
            else
            {
                var nextBraceIndex = FindNextBraceIndex(text, i);
                if (nextBraceIndex == -1)
                {
                    builder.Write(text[i..]);
                    break;
                }

                builder.Write(text[i..nextBraceIndex]);
                i = nextBraceIndex;
            }
    }

    /// <summary>
    ///     Finds the index of the next brace character.
    /// </summary>
    /// <param name="text">The text to search.</param>
    /// <param name="startIndex">The starting index.</param>
    /// <returns>The index of the next brace, or -1 if not found.</returns>
    private static int FindNextBraceIndex(string text, int startIndex)
    {
        for (var i = startIndex; i < text.Length; i++)
            if (text[i] == '{' || text[i] == '}')
                return i;

        return -1;
    }

    /// <summary>
    ///     Inserts a field by its code name.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="fieldCode">The field code (e.g., PAGE, DATE).</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the code is not one of the supported fields.
    /// </exception>
    private static void InsertFieldByCode(DocumentBuilder builder, string fieldCode)
    {
        var code = fieldCode.ToUpper();

        // Anything not in the map used to be inserted verbatim as a live field, which made header
        // text a way to run field codes the caller was never offered. Measured against the pinned
        // Aspose.Words: an INCLUDETEXT field does not resolve on save or PDF render, but it does
        // resolve on Document.UpdateFields(), which this server calls in several handlers -
        // including the one that reads headers back. That turned "set a header, then read the
        // headers" into an arbitrary local file read, with the path inside a field code where the
        // allowlist never sees it. Only the documented fields are accepted now.
        if (!FieldCodeMap.TryGetValue(code, out var fieldType))
            throw new ArgumentException(
                $"Unsupported field code '{fieldCode}'. Supported: "
                + string.Join(", ", FieldCodeMap.Keys.Order())
                + ". Use '{{' and '}}' for a literal brace.");

        builder.InsertField(fieldType, true);
    }

    /// <summary>
    ///     Applies font settings to the document builder.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="fontSettings">The font settings to apply.</param>
    private static void ApplyFontSettings(DocumentBuilder builder, FontSettings fontSettings)
    {
        if (!string.IsNullOrEmpty(fontSettings.FontName))
            builder.Font.Name = fontSettings.FontName;

        if (!string.IsNullOrEmpty(fontSettings.FontNameAscii))
            builder.Font.NameAscii = fontSettings.FontNameAscii;

        if (!string.IsNullOrEmpty(fontSettings.FontNameFarEast))
            builder.Font.NameFarEast = fontSettings.FontNameFarEast;

        if (fontSettings.FontSize.HasValue)
            builder.Font.Size = fontSettings.FontSize.Value;
    }
}
