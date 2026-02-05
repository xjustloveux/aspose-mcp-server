using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Toc;
using AsposeMcpServer.Results.Pdf.Toc;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Toc;

/// <summary>
///     Tests for <see cref="GetPdfTocHandler" />.
///     Validates TOC retrieval from documents with and without outlines and read-only behavior.
/// </summary>
public class GetPdfTocHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfTocHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotMarkModified()
    {
        var doc = CreateDocumentWithOutlines(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Entry Properties

    [Fact]
    public void Execute_ReturnsEntryProperties()
    {
        var doc = CreateDocumentWithOutlines(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTocPdfResult>(res);
        Assert.True(result.Entries.Count > 0);
        var entry = result.Entries[0];
        Assert.False(string.IsNullOrEmpty(entry.Title));
        Assert.True(entry.Level >= 1);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_OnDocumentWithoutToc_ReturnsHasTocFalse()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTocPdfResult>(res);
        Assert.False(result.HasToc);
        Assert.Equal(0, result.EntryCount);
        Assert.Contains("No table of contents found", result.Message);
    }

    [Fact]
    public void Execute_OnDocumentWithOutlines_ReturnsEntries()
    {
        var doc = CreateDocumentWithOutlines(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTocPdfResult>(res);
        Assert.True(result.HasToc);
        Assert.Equal(3, result.EntryCount);
        Assert.Equal(3, result.Entries.Count);
    }

    [Fact]
    public void Execute_ReturnsCorrectEntryCount()
    {
        var doc = CreateDocumentWithOutlines(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTocPdfResult>(res);
        Assert.Equal(3, result.EntryCount);
    }

    [Fact]
    public void Execute_OnEmptyDocument_ReturnsNoToc()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTocPdfResult>(res);
        Assert.False(result.HasToc);
        Assert.Equal(0, result.EntryCount);
        Assert.Empty(result.Entries);
    }

    [Fact]
    public void Execute_OnDocumentWithTocPage_ReturnsHasTocTrue()
    {
        var doc = CreateDocumentWithTocPage();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTocPdfResult>(res);
        Assert.True(result.HasToc);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithOutlines(int count)
    {
        var doc = new Document();
        for (var i = 0; i < count + 1; i++)
            doc.Pages.Add();

        for (var i = 0; i < count; i++)
        {
            var outline = new OutlineItemCollection(doc.Outlines)
            {
                Title = $"Chapter {i + 1}",
                Destination = new GoToAction(doc.Pages[i + 1])
            };
            doc.Outlines.Add(outline);
        }

        return doc;
    }

    private static Document CreateDocumentWithTocPage()
    {
        var doc = new Document();
        var tocPage = doc.Pages.Add();
        tocPage.TocInfo = new TocInfo
        {
            Title = new TextFragment("Table of Contents")
        };
        doc.Pages.Add();
        return doc;
    }

    #endregion
}
