using Aspose.Words;
using Aspose.Words.Fields;
using WordParagraph = Aspose.Words.Paragraph;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

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
    public static Aspose.Words.HeaderFooter GetOrCreateHeaderFooter(Section section, Document doc,
        HeaderFooterType hfType)
    {
        var headerFooter = section.HeadersFooters[hfType];
        if (headerFooter == null)
        {
            headerFooter = new Aspose.Words.HeaderFooter(doc, hfType);
            section.HeadersFooters.Add(headerFooter);
        }

        return headerFooter;
    }

    /// <summary>
    ///     Clears only the text content from a header/footer, preserving other elements.
    /// </summary>
    /// <param name="headerFooter">The header/footer to clear text from.</param>
    public static void ClearTextOnly(Aspose.Words.HeaderFooter headerFooter)
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
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="text">The text to insert. Field codes start with { and end with }.</param>
    /// <param name="fontName">Optional font name.</param>
    /// <param name="fontNameAscii">Optional ASCII font name.</param>
    /// <param name="fontNameFarEast">Optional Far East font name.</param>
    /// <param name="fontSize">Optional font size.</param>
    public static void InsertTextOrField(DocumentBuilder builder, string text, string? fontName,
        string? fontNameAscii, string? fontNameFarEast, double? fontSize)
    {
        ApplyFontSettings(builder, fontName, fontNameAscii, fontNameFarEast, fontSize);

        if (IsFieldCode(text))
            InsertField(builder, text);
        else
            builder.Write(text);
    }

    /// <summary>
    ///     Applies font settings to the document builder.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="fontName">Optional font name.</param>
    /// <param name="fontNameAscii">Optional ASCII font name.</param>
    /// <param name="fontNameFarEast">Optional Far East font name.</param>
    /// <param name="fontSize">Optional font size.</param>
    private static void ApplyFontSettings(DocumentBuilder builder, string? fontName, string? fontNameAscii,
        string? fontNameFarEast, double? fontSize)
    {
        if (!string.IsNullOrEmpty(fontName))
            builder.Font.Name = fontName;

        if (!string.IsNullOrEmpty(fontNameAscii))
            builder.Font.NameAscii = fontNameAscii;

        if (!string.IsNullOrEmpty(fontNameFarEast))
            builder.Font.NameFarEast = fontNameFarEast;

        if (fontSize.HasValue)
            builder.Font.Size = fontSize.Value;
    }

    /// <summary>
    ///     Checks if the text is a field code (enclosed in braces).
    /// </summary>
    /// <param name="text">The text to check.</param>
    /// <returns>True if the text is a field code.</returns>
    private static bool IsFieldCode(string text)
    {
        return text.StartsWith('{') && text.EndsWith('}');
    }

    /// <summary>
    ///     Inserts a field based on the field code.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="fieldCode">The field code (e.g., {PAGE}, {DATE}).</param>
    private static void InsertField(DocumentBuilder builder, string fieldCode)
    {
        var code = fieldCode.Trim('{', '}', ' ').ToUpper();

        if (FieldCodeMap.TryGetValue(code, out var fieldType))
            builder.InsertField(fieldType, true);
        else
            builder.InsertField($" {code} ", null);
    }
}
