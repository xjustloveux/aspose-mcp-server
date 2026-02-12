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
    ///     Inserts text or a field code (like PAGE, DATE) into the document.
    ///     Supports mixed content like "Page {PAGE} of {NUMPAGES}".
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
    private static void InsertFieldByCode(DocumentBuilder builder, string fieldCode)
    {
        var code = fieldCode.ToUpper();

        if (FieldCodeMap.TryGetValue(code, out var fieldType))
            builder.InsertField(fieldType, true);
        else
            builder.InsertField($" {code} ", null);
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
