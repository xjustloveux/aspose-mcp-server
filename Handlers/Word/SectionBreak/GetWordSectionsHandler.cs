using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.SectionBreak;

/// <summary>
///     Handler for getting section information from Word documents.
/// </summary>
public class GetWordSectionsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets section information from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sectionIndex (null for all sections)
    /// </param>
    /// <returns>JSON string containing section information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetWordSectionsParameters(parameters);

        var doc = context.Document;

        if (p.SectionIndex.HasValue)
        {
            if (p.SectionIndex.Value < 0 || p.SectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException(
                    $"sectionIndex must be between 0 and {doc.Sections.Count - 1}, got: {p.SectionIndex.Value}");

            var section = doc.Sections[p.SectionIndex.Value];
            var sectionInfo = WordSectionHelper.BuildSectionInfo(section, p.SectionIndex.Value);

            var result = new
            {
                totalSections = doc.Sections.Count,
                section = sectionInfo
            };

            return JsonResult(result);
        }
        else
        {
            List<object> sectionList = [];
            for (var i = 0; i < doc.Sections.Count; i++)
            {
                var section = doc.Sections[i];
                sectionList.Add(WordSectionHelper.BuildSectionInfo(section, i));
            }

            var result = new
            {
                totalSections = doc.Sections.Count,
                sections = sectionList
            };

            return JsonResult(result);
        }
    }

    private static GetWordSectionsParameters ExtractGetWordSectionsParameters(OperationParameters parameters)
    {
        return new GetWordSectionsParameters(
            parameters.GetOptional<int?>("sectionIndex"));
    }

    private sealed record GetWordSectionsParameters(int? SectionIndex);
}
