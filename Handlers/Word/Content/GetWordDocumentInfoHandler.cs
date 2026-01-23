using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Content;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Content;

/// <summary>
///     Handler for getting Word document metadata and properties.
/// </summary>
[ResultType(typeof(GetWordDocumentInfoResult))]
public class GetWordDocumentInfoHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_document_info";

    /// <summary>
    ///     Gets document metadata and properties as JSON.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: includeTabStops (default: false)
    /// </param>
    /// <returns>JSON string containing document metadata and properties.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetDocumentInfoParameters(parameters);

        var document = context.Document;
        var props = document.BuiltInDocumentProperties;

        List<TabStopsByParagraph>? tabStopsList = null;
        if (p.IncludeTabStops)
        {
            tabStopsList = [];
            var sectionIndex = 0;
            foreach (var section in document.Sections.Cast<Section>())
            {
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
                for (var paraIndex = 0; paraIndex < paragraphs.Count; paraIndex++)
                {
                    var para = paragraphs[paraIndex];
                    if (para.ParagraphFormat.TabStops.Count > 0)
                    {
                        List<TabStopInfo> stops = [];
                        for (var i = 0; i < para.ParagraphFormat.TabStops.Count; i++)
                        {
                            var tabStop = para.ParagraphFormat.TabStops[i];
                            stops.Add(new TabStopInfo
                            {
                                Position = tabStop.Position,
                                Alignment = tabStop.Alignment.ToString()
                            });
                        }

                        tabStopsList.Add(new TabStopsByParagraph
                        {
                            SectionIndex = sectionIndex,
                            ParagraphIndex = paraIndex,
                            TabStops = stops
                        });
                    }
                }

                sectionIndex++;
            }
        }

        return new GetWordDocumentInfoResult
        {
            Title = props.Title,
            Author = props.Author,
            Subject = props.Subject,
            Created = props.CreatedTime.ToString("yyyy-MM-dd HH:mm:ss"),
            Modified = props.LastSavedTime.ToString("yyyy-MM-dd HH:mm:ss"),
            Pages = props.Pages,
            Sections = document.Sections.Count,
            TabStopsIncluded = p.IncludeTabStops,
            TabStops = tabStopsList
        };
    }

    /// <summary>
    ///     Extracts parameters for the get document info operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetDocumentInfoParameters ExtractGetDocumentInfoParameters(OperationParameters parameters)
    {
        var includeTabStops = parameters.GetOptional("includeTabStops", false);

        return new GetDocumentInfoParameters(includeTabStops);
    }

    /// <summary>
    ///     Parameters for the get document info operation.
    /// </summary>
    /// <param name="IncludeTabStops">Whether to include tab stops information.</param>
    private sealed record GetDocumentInfoParameters(bool IncludeTabStops);
}
