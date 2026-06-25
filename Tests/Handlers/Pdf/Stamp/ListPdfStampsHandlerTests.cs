using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Stamp;
using AsposeMcpServer.Results.Pdf.Stamp;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Stamp;

/// <summary>
///     Tests for <see cref="ListPdfStampsHandler" />.
///     Validates stamp listing with page filtering and read-only behavior.
/// </summary>
public class ListPdfStampsHandlerTests : PdfHandlerTestBase
{
    private readonly ListPdfStampsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_List()
    {
        Assert.Equal("list", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotMarkModified()
    {
        var doc = CreateDocumentWithStampAnnotations(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Stamp Properties

    [Fact]
    public void Execute_ReturnsStampProperties()
    {
        var doc = CreateDocumentWithStampAnnotations(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStampsPdfResult>(res);
        Assert.True(result.Stamps.Count > 0);
        var stamp = result.Stamps[0];
        Assert.True(stamp.PageIndex > 0);
        Assert.True(stamp.Index > 0);
        Assert.Equal("stamp", stamp.Type);
    }

    [Fact]
    public void Execute_WithNonStampAnnotationBeforeStamp_ReportsStampRelativeIndex()
    {
        // Adversarial: a non-stamp annotation precedes the stamp in page.Annotations. The reported Index
        // must be stamp-relative (1 = first stamp) so it can be fed straight into remove's stampIndex,
        // not the absolute page.Annotations position (which would be 2 here).
        var doc = new Document();
        var page = doc.Pages.Add();
        page.Annotations.Add(new TextAnnotation(page, new Rectangle(10, 10, 60, 60))
        {
            Title = "note", Contents = "a note"
        });
        page.Annotations.Add(new StampAnnotation(page, new Rectangle(100, 100, 200, 200))
        {
            Contents = "Stamp 1"
        });

        var res = _handler.Execute(CreateContext(doc), CreateEmptyParameters());

        var result = Assert.IsType<GetStampsPdfResult>(res);
        var stamp = Assert.Single(result.Stamps);
        Assert.Equal(1, stamp.Index);
    }

    #endregion

    #region Basic List Operations

    [Fact]
    public void Execute_OnEmptyDocument_ReturnsEmptyList()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStampsPdfResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Empty(result.Stamps);
    }

    [Fact]
    public void Execute_WithStampAnnotations_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithStampAnnotations(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStampsPdfResult>(res);
        Assert.Equal(2, result.Count);
        Assert.Equal(2, result.Stamps.Count);
    }

    [Fact]
    public void Execute_WithPageIndexZero_ReturnsAllPages()
    {
        var doc = CreateDocumentWithPages(2);
        AddStampAnnotationToPage(doc, 1);
        AddStampAnnotationToPage(doc, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStampsPdfResult>(res);
        Assert.Equal(2, result.Count);
        Assert.Contains("all pages", result.Message);
    }

    [Fact]
    public void Execute_WithSpecificPageIndex_ReturnsPageStamps()
    {
        var doc = CreateDocumentWithPages(2);
        AddStampAnnotationToPage(doc, 1);
        AddStampAnnotationToPage(doc, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetStampsPdfResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Contains("page 1", result.Message);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithStampAnnotations(int count)
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        for (var i = 0; i < count; i++)
            AddStampAnnotation(page);
        return doc;
    }

    private static void AddStampAnnotationToPage(Document doc, int pageIndex)
    {
        AddStampAnnotation(doc.Pages[pageIndex]);
    }

    private static void AddStampAnnotation(Aspose.Pdf.Page page)
    {
        var annotation = new StampAnnotation(page, new Rectangle(100, 100, 200, 200))
        {
            Contents = "Test Stamp"
        };
        page.Annotations.Add(annotation);
    }

    #endregion
}
