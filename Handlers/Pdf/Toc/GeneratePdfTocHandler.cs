using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Toc;

/// <summary>
///     Handler for generating a table of contents page in PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class GeneratePdfTocHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "generate";

    /// <summary>
    ///     Generates a table of contents page and inserts it into the PDF document.
    ///     If the document has outlines (bookmarks), uses them to build TOC entries.
    ///     Otherwise, generates entries from each page in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: title (default "Table of Contents"), depth (default 3), tocPage (default 1)
    /// </param>
    /// <returns>Success message with TOC generation details.</returns>
    /// <exception cref="ArgumentException">Thrown when tocPage is out of valid range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var generateParams = ExtractGenerateParameters(parameters);

        var document = context.Document;

        if (generateParams.TocPage < 1 || generateParams.TocPage > document.Pages.Count + 1)
            throw new ArgumentException(
                $"tocPage must be between 1 and {document.Pages.Count + 1}");

        var tocPageObj = document.Pages.Insert(generateParams.TocPage);

        var titleFragment = new TextFragment(generateParams.Title)
        {
            TextState = { FontSize = 16, FontStyle = FontStyles.Bold }
        };

        tocPageObj.TocInfo = new TocInfo { Title = titleFragment };

        var entryCount = 0;
        if (document.Outlines.Count > 0)
            foreach (var outline in document.Outlines)
            {
                if (outline.Level <= generateParams.Depth)
                {
                    AddTocHeading(tocPageObj, outline);
                    entryCount++;
                }

                AddChildOutlines(tocPageObj, outline, generateParams.Depth, ref entryCount);
            }
        else
            for (var i = 1; i <= document.Pages.Count; i++)
            {
                if (i == generateParams.TocPage) continue;

                var heading = new Heading(1)
                {
                    IsInList = true,
                    DestinationPage = document.Pages[i]
                };

                var segment = new TextSegment($"Page {i}");
                heading.Segments.Add(segment);
                tocPageObj.Paragraphs.Add(heading);
                entryCount++;
            }

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Generated TOC with {entryCount} entries at page {generateParams.TocPage}."
        };
    }

    /// <summary>
    ///     Recursively processes child outlines and adds them as TOC headings.
    /// </summary>
    /// <param name="tocPage">The TOC page to add headings to.</param>
    /// <param name="parent">The parent outline item collection.</param>
    /// <param name="maxDepth">The maximum heading depth to include.</param>
    /// <param name="count">Reference counter for the number of entries added.</param>
    private static void AddChildOutlines(
        Aspose.Pdf.Page tocPage,
        OutlineItemCollection parent,
        int maxDepth,
        ref int count)
    {
        foreach (var child in parent)
        {
            if (child.Level > maxDepth) continue;

            AddTocHeading(tocPage, child);
            count++;

            AddChildOutlines(tocPage, child, maxDepth, ref count);
        }
    }

    /// <summary>
    ///     Adds a single TOC heading entry from an outline item.
    /// </summary>
    /// <param name="tocPage">The TOC page to add the heading to.</param>
    /// <param name="outline">The outline item to create the heading from.</param>
    private static void AddTocHeading(Aspose.Pdf.Page tocPage, OutlineItemCollection outline)
    {
        var heading = new Heading(outline.Level)
        {
            IsInList = true,
            DestinationPage = ExtractDestinationPage(outline)
        };

        var segment = new TextSegment(outline.Title ?? string.Empty);
        heading.Segments.Add(segment);
        tocPage.Paragraphs.Add(heading);
    }

    /// <summary>
    ///     Extracts the destination page from an outline item.
    /// </summary>
    /// <param name="outline">The outline item to extract the destination from.</param>
    /// <returns>The destination page, or null if not available.</returns>
    private static Aspose.Pdf.Page? ExtractDestinationPage(OutlineItemCollection outline)
    {
        Aspose.Pdf.Page? page = null;

        if (outline.Destination is ExplicitDestination explicitDest)
            page = explicitDest.Page;
        else if (outline.Action is GoToAction { Destination: ExplicitDestination actionDest })
            page = actionDest.Page;

        return page;
    }

    /// <summary>
    ///     Extracts generate parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted generate parameters.</returns>
    private static GenerateParameters ExtractGenerateParameters(OperationParameters parameters)
    {
        return new GenerateParameters(
            parameters.GetOptional("title", "Table of Contents"),
            parameters.GetOptional("depth", 3),
            parameters.GetOptional("tocPage", 1)
        );
    }

    /// <summary>
    ///     Record to hold generate TOC parameters.
    /// </summary>
    private sealed record GenerateParameters(string Title, int Depth, int TocPage);
}
