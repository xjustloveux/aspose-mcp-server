using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Toc;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Toc;

/// <summary>
///     Tests for <see cref="RemovePdfTocHandler" />.
///     Validates TOC page removal and modification tracking behavior.
/// </summary>
public class RemovePdfTocHandlerTests : PdfHandlerTestBase
{
    private readonly RemovePdfTocHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Remove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    #endregion

    #region Page Count Verification

    [Fact]
    public void Execute_PreservesNonTocPages()
    {
        var doc = CreateDocumentWithTocPage(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(3, doc.Pages.Count);
        for (var i = 1; i <= doc.Pages.Count; i++)
            Assert.Null(doc.Pages[i].TocInfo);
    }

    #endregion

    #region Basic Remove Operations

    [Fact]
    public void Execute_RemovesTocPages()
    {
        var doc = CreateDocumentWithTocPage(3);
        var initialPageCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Removed 1 TOC page(s)", result.Message);
        Assert.Equal(initialPageCount - 1, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_OnDocumentWithoutToc_ReturnsAppropriateMessage()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("No TOC pages found", result.Message);
    }

    [Fact]
    public void Execute_RemovesMultipleTocPages()
    {
        var doc = CreateDocumentWithMultipleTocPages(2, 2);
        var initialPageCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Removed 2 TOC page(s)", result.Message);
        Assert.Equal(initialPageCount - 2, doc.Pages.Count);
        AssertModified(context);
    }

    #endregion

    #region Modification Tracking

    [Fact]
    public void Execute_MarksModifiedWhenPagesRemoved()
    {
        var doc = CreateDocumentWithTocPage(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_DoesNotMarkModifiedWhenNoTocFound()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTocPage(int contentPageCount)
    {
        var doc = new Document();
        var tocPage = doc.Pages.Add();
        tocPage.TocInfo = new TocInfo
        {
            Title = new TextFragment("Table of Contents")
        };
        for (var i = 0; i < contentPageCount; i++)
            doc.Pages.Add();
        return doc;
    }

    private static Document CreateDocumentWithMultipleTocPages(int tocPageCount, int contentPageCount)
    {
        var doc = new Document();
        for (var i = 0; i < tocPageCount; i++)
        {
            var tocPage = doc.Pages.Add();
            tocPage.TocInfo = new TocInfo
            {
                Title = new TextFragment($"TOC Part {i + 1}")
            };
        }

        for (var i = 0; i < contentPageCount; i++)
            doc.Pages.Add();
        return doc;
    }

    #endregion
}
