using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Toc;

namespace AsposeMcpServer.Handlers.Pdf.Toc;

/// <summary>
///     Handler for getting table of contents information from PDF documents.
/// </summary>
[ResultType(typeof(GetTocPdfResult))]
public class GetPdfTocHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Retrieves table of contents information from the PDF document.
    ///     Checks for pages with TocInfo and collects entries from document outlines.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>TOC information including entries and their page targets.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;

        var hasToc = false;
        var entries = new List<TocEntryPdfInfo>();

        for (var i = 1; i <= document.Pages.Count; i++)
        {
            if (document.Pages[i].TocInfo == null) continue;

            hasToc = true;
            break;
        }

        if (document.Outlines.Count > 0)
        {
            hasToc = true;
            CollectOutlineEntries(document, document.Outlines, entries, 1);
        }

        return new GetTocPdfResult
        {
            HasToc = hasToc,
            EntryCount = entries.Count,
            Entries = entries,
            Message = hasToc
                ? $"Found TOC with {entries.Count} entries."
                : "No table of contents found in document."
        };
    }

    /// <summary>
    ///     Collects TOC entries from the top-level outline collection.
    /// </summary>
    /// <param name="document">The PDF document.</param>
    /// <param name="outlines">The outline collection to process.</param>
    /// <param name="entries">The list to add entries to.</param>
    /// <param name="level">The current heading level.</param>
    private static void CollectOutlineEntries(
        Document document,
        OutlineCollection outlines,
        List<TocEntryPdfInfo> entries,
        int level)
    {
        foreach (var outline in outlines)
        {
            var pageNumber = ExtractPageNumber(document, outline);

            entries.Add(new TocEntryPdfInfo
            {
                Title = outline.Title ?? string.Empty,
                PageNumber = pageNumber,
                Level = level
            });

            if (outline.Count > 0)
                CollectChildOutlineEntries(document, outline, entries, level + 1);
        }
    }

    /// <summary>
    ///     Collects TOC entries from child outline items recursively.
    /// </summary>
    /// <param name="document">The PDF document.</param>
    /// <param name="parent">The parent outline item collection.</param>
    /// <param name="entries">The list to add entries to.</param>
    /// <param name="level">The current heading level.</param>
    private static void CollectChildOutlineEntries(
        Document document,
        OutlineItemCollection parent,
        List<TocEntryPdfInfo> entries,
        int level)
    {
        foreach (var child in parent)
        {
            var pageNumber = ExtractPageNumber(document, child);

            entries.Add(new TocEntryPdfInfo
            {
                Title = child.Title ?? string.Empty,
                PageNumber = pageNumber,
                Level = level
            });

            if (child.Count > 0)
                CollectChildOutlineEntries(document, child, entries, level + 1);
        }
    }

    /// <summary>
    ///     Extracts the target page number from an outline item.
    /// </summary>
    /// <param name="document">The PDF document.</param>
    /// <param name="outline">The outline item to extract the page number from.</param>
    /// <returns>The 1-based page number, or 0 if not available.</returns>
    private static int ExtractPageNumber(Document document, OutlineItemCollection outline)
    {
        Aspose.Pdf.Page? page = null;

        if (outline.Destination is ExplicitDestination explicitDest)
            page = explicitDest.Page;
        else if (outline.Action is GoToAction { Destination: ExplicitDestination actionDest })
            page = actionDest.Page;

        if (page == null)
            return 0;

        var pageIndex = document.Pages.IndexOf(page);
        return pageIndex > 0 ? pageIndex : 0;
    }
}
