using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Toc;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Toc;

/// <summary>
///     Covers TEST-09. TOC generation has two paths: build entries from the document outlines, or,
///     when there are none, list one entry per page. Only the page-listing path had tests, so the
///     outline path — the one that runs on any real document with bookmarks — was unverified,
///     including the depth limit and the recursion into child outlines.
/// </summary>
public class GeneratePdfTocFromOutlinesTests : PdfHandlerTestBase
{
    private readonly GeneratePdfTocHandler _handler = new();

    /// <summary>
    ///     Builds a document with a three-level outline tree: two top-level entries, each with one
    ///     child, and one grandchild under the first child.
    /// </summary>
    /// <returns>The document, with four pages.</returns>
    private static Document CreateDocumentWithNestedOutlines()
    {
        var document = new Document();
        for (var i = 0; i < 4; i++) document.Pages.Add();

        var first = new OutlineItemCollection(document.Outlines)
        {
            Title = "Chapter 1",
            Action = new GoToAction(document.Pages[1])
        };
        document.Outlines.Add(first);

        var firstChild = new OutlineItemCollection(document.Outlines)
        {
            Title = "Section 1.1",
            Action = new GoToAction(document.Pages[2])
        };
        first.Add(firstChild);

        var grandchild = new OutlineItemCollection(document.Outlines)
        {
            Title = "Section 1.1.1",
            Action = new GoToAction(document.Pages[3])
        };
        firstChild.Add(grandchild);

        var second = new OutlineItemCollection(document.Outlines)
        {
            Title = "Chapter 2",
            Action = new GoToAction(document.Pages[4])
        };
        document.Outlines.Add(second);

        return document;
    }

    [Fact]
    public void Execute_WithOutlines_ShouldUseThemInsteadOfListingPages()
    {
        var document = CreateDocumentWithNestedOutlines();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "depth", 3 }
        });

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        // Four outline items exist across three levels; listing pages instead would report five,
        // one per page plus the inserted TOC page being skipped.
        Assert.Contains("4 entries", result.Message);
    }

    [Fact]
    public void Execute_WithDepthLimit_ShouldSkipDeeperOutlines()
    {
        var document = CreateDocumentWithNestedOutlines();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "depth", 2 }
        });

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        // Level 3 ("Section 1.1.1") is excluded, leaving the two chapters and one section.
        Assert.Contains("3 entries", result.Message);
    }

    [Fact]
    public void Execute_WithDepthOne_ShouldKeepOnlyTopLevelOutlines()
    {
        var document = CreateDocumentWithNestedOutlines();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "depth", 1 }
        });

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        Assert.Contains("2 entries", result.Message);
    }

    [SkippableFact]
    public void Execute_WithOutlines_ShouldWriteAReadableTocPage()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Evaluation mode truncates the rendered text");
        var document = CreateDocumentWithNestedOutlines();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Contents" },
            { "depth", 3 },
            { "tocPage", 1 }
        });

        _handler.Execute(context, parameters);

        var path = CreateTestFilePath("toc_outlines.pdf");
        document.Save(path);

        using var saved = new Document(path);
        var absorber = new TextFragmentAbsorber();
        saved.Pages[1].Accept(absorber);
        var text = string.Join(" ", absorber.TextFragments.Select(f => f.Text));

        Assert.Contains("Contents", text);
        Assert.Contains("Chapter 1", text);
        Assert.Contains("Section 1.1", text);
        Assert.Contains("Chapter 2", text);
    }

    [Fact]
    public void Execute_WithOutlines_ShouldInsertTheTocAtTheRequestedPage()
    {
        var document = CreateDocumentWithNestedOutlines();
        var pagesBefore = document.Pages.Count;
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tocPage", 2 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(pagesBefore + 1, document.Pages.Count);
        Assert.NotNull(document.Pages[2].TocInfo);
    }
}
