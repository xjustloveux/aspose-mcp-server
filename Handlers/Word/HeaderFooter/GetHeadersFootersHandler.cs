using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using Section = Aspose.Words.Section;

namespace AsposeMcpServer.Handlers.Word.HeaderFooter;

public class GetHeadersFootersHandler : OperationHandlerBase<Document>
{
    public override string Operation => "get";

    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        doc.UpdateFields();

        var sections = sectionIndex.HasValue && sectionIndex.Value != -1
            ? [doc.Sections[sectionIndex.Value]]
            : doc.Sections.Cast<Section>().ToArray();

        if (sectionIndex.HasValue && sectionIndex.Value != -1 &&
            (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count))
            throw new ArgumentException(
                $"Section index {sectionIndex.Value} is out of range (document has {doc.Sections.Count} sections)");

        List<object> sectionsList = [];

        for (var i = 0; i < sections.Length; i++)
        {
            var section = sections[i];
            var actualIndex = sectionIndex.HasValue && sectionIndex.Value != -1 ? sectionIndex.Value : i;

            var headerTypes = new[]
            {
                (HeaderFooterType.HeaderPrimary, "primary"),
                (HeaderFooterType.HeaderFirst, "firstPage"),
                (HeaderFooterType.HeaderEven, "evenPage")
            };

            var headers = new Dictionary<string, string?>();
            foreach (var (type, name) in headerTypes)
            {
                var header = section.HeadersFooters[type];
                if (header != null)
                {
                    var headerText = header.ToString(SaveFormat.Text).Trim();
                    if (!string.IsNullOrEmpty(headerText))
                        headers[name] = headerText;
                }
            }

            var footerTypes = new[]
            {
                (HeaderFooterType.FooterPrimary, "primary"),
                (HeaderFooterType.FooterFirst, "firstPage"),
                (HeaderFooterType.FooterEven, "evenPage")
            };

            var footers = new Dictionary<string, string?>();
            foreach (var (type, name) in footerTypes)
            {
                var footer = section.HeadersFooters[type];
                if (footer != null)
                {
                    var footerText = footer.ToString(SaveFormat.Text).Trim();
                    if (!string.IsNullOrEmpty(footerText))
                        footers[name] = footerText;
                }
            }

            sectionsList.Add(new
            {
                sectionIndex = actualIndex,
                headers = headers.Count > 0 ? headers : null,
                footers = footers.Count > 0 ? footers : null
            });
        }

        var result = new
        {
            totalSections = doc.Sections.Count,
            queriedSectionIndex = sectionIndex,
            sections = sectionsList
        };

        return JsonResult(result);
    }
}
