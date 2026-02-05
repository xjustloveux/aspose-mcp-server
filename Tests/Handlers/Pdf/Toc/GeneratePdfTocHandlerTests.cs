using AsposeMcpServer.Handlers.Pdf.Toc;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Toc;

/// <summary>
///     Tests for <see cref="GeneratePdfTocHandler" />.
///     Validates TOC generation with various parameters and error handling.
/// </summary>
public class GeneratePdfTocHandlerTests : PdfHandlerTestBase
{
    private readonly GeneratePdfTocHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Generate()
    {
        Assert.Equal("generate", _handler.Operation);
    }

    #endregion

    #region Modification Tracking

    [Fact]
    public void Execute_MarksContextAsModified()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Basic Generate Operations

    [Fact]
    public void Execute_GeneratesTocWithDefaultSettings()
    {
        var doc = CreateDocumentWithPages(3);
        var initialPageCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Generated TOC", result.Message);
        Assert.Equal(initialPageCount + 1, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_GeneratesTocWithCustomTitle()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Contents" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Generated TOC", result.Message);
        Assert.NotNull(doc.Pages[1].TocInfo);
        AssertModified(context);
    }

    [Fact]
    public void Execute_GeneratesTocAtSpecifiedPagePosition()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tocPage", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("page 2", result.Message);
        Assert.NotNull(doc.Pages[2].TocInfo);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithDepthParameter_GeneratesToc()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "depth", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Generated TOC", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_OnMultiPageDocument_CreatesEntries()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("entries", result.Message);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Execute_WithInvalidTocPage_ThrowsArgumentException(int invalidTocPage)
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tocPage", invalidTocPage }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithTocPageBeyondRange_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tocPage", 100 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Page Insertion

    [Fact]
    public void Execute_InsertsNewPageForToc()
    {
        var doc = CreateDocumentWithPages(3);
        var initialPageCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialPageCount + 1, doc.Pages.Count);
    }

    [Fact]
    public void Execute_TocPageHasTocInfo()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tocPage", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.NotNull(doc.Pages[1].TocInfo);
    }

    #endregion
}
