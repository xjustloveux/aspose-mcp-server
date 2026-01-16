using Aspose.Words;

namespace AsposeMcpServer.Handlers.Word.SectionBreak;

/// <summary>
///     Helper class for Word section operations.
/// </summary>
public static class WordSectionHelper
{
    /// <summary>
    ///     Converts section break type string to SectionStart enum.
    /// </summary>
    /// <param name="sectionBreakType">The section break type string.</param>
    /// <returns>The corresponding SectionStart enum value.</returns>
    public static SectionStart GetSectionStart(string sectionBreakType)
    {
        return sectionBreakType switch
        {
            "NextPage" => SectionStart.NewPage,
            "Continuous" => SectionStart.Continuous,
            "EvenPage" => SectionStart.EvenPage,
            "OddPage" => SectionStart.OddPage,
            _ => SectionStart.NewPage
        };
    }

    /// <summary>
    ///     Converts SectionStart enum to human-readable name.
    /// </summary>
    /// <param name="sectionStart">The SectionStart enum value.</param>
    /// <returns>Human-readable section break type name.</returns>
    public static string GetSectionStartName(SectionStart sectionStart)
    {
        return sectionStart switch
        {
            SectionStart.NewPage => "NextPage",
            SectionStart.Continuous => "Continuous",
            SectionStart.EvenPage => "EvenPage",
            SectionStart.OddPage => "OddPage",
            SectionStart.NewColumn => "NewColumn",
            _ => sectionStart.ToString()
        };
    }

    /// <summary>
    ///     Builds section information as a structured object.
    /// </summary>
    /// <param name="section">The section to extract information from.</param>
    /// <param name="index">The index of the section in the document.</param>
    /// <returns>An object containing section information.</returns>
    public static object BuildSectionInfo(Section section, int index)
    {
        var pageSetup = section.PageSetup;

        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true);
        var tables = section.Body.GetChildNodes(NodeType.Table, true);
        var shapes = section.Body.GetChildNodes(NodeType.Shape, true);

        var headerCount = 0;
        var footerCount = 0;
        foreach (var hf in section.HeadersFooters.Cast<Aspose.Words.HeaderFooter>())
            if (hf.HeaderFooterType is HeaderFooterType.HeaderPrimary or HeaderFooterType.HeaderFirst
                or HeaderFooterType.HeaderEven)
            {
                if (!string.IsNullOrWhiteSpace(hf.GetText()))
                    headerCount++;
            }
            else if (hf.HeaderFooterType is HeaderFooterType.FooterPrimary or HeaderFooterType.FooterFirst
                         or HeaderFooterType.FooterEven && !string.IsNullOrWhiteSpace(hf.GetText()))
            {
                footerCount++;
            }

        return new
        {
            index,
            sectionBreak = new
            {
                type = GetSectionStartName(pageSetup.SectionStart)
            },
            pageSetup = new
            {
                paperSize = pageSetup.PaperSize.ToString(),
                orientation = pageSetup.Orientation.ToString(),
                margins = new
                {
                    top = pageSetup.TopMargin,
                    bottom = pageSetup.BottomMargin,
                    left = pageSetup.LeftMargin,
                    right = pageSetup.RightMargin
                },
                headerFooterDistance = new
                {
                    header = pageSetup.HeaderDistance,
                    footer = pageSetup.FooterDistance
                },
                pageNumberStart = pageSetup.RestartPageNumbering ? pageSetup.PageStartingNumber : (int?)null,
                differentFirstPage = pageSetup.DifferentFirstPageHeaderFooter,
                differentOddEvenPages = pageSetup.OddAndEvenPagesHeaderFooter,
                columnCount = pageSetup.TextColumns.Count
            },
            contentStatistics = new
            {
                paragraphs = paragraphs.Count,
                tables = tables.Count,
                shapes = shapes.Count
            },
            headersFooters = new
            {
                headers = headerCount,
                footers = footerCount
            }
        };
    }
}
