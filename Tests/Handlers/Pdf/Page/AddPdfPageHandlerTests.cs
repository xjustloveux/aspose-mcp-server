using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Page;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Page;

public class AddPdfPageHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfPageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Count and InsertAt Combined

    [SkippableTheory]
    [InlineData(2, 1, false)]
    [InlineData(3, 2, true)]
    [InlineData(5, 1, true)]
    public void Execute_WithCountAndInsertAt_InsertsMultiplePagesAtPosition(int count, int insertAt,
        bool requiresLicense)
    {
        if (requiresLicense)
            SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", count },
            { "insertAt", insertAt }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(count.ToString(), result.Message);
        Assert.Equal(3 + count, doc.Pages.Count);
        AssertModified(context);
    }

    #endregion

    #region Basic Page Addition

    [Fact]
    public void Execute_AddsPageToDocument()
    {
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Equal(initialCount + 1, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_MultipleTimes_AddsMultiplePages()
    {
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);
        _handler.Execute(context, parameters);
        _handler.Execute(context, parameters);

        Assert.Equal(initialCount + 3, doc.Pages.Count);
    }

    [Fact]
    public void Execute_DefaultCount_AddsOnePage()
    {
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount + 1, doc.Pages.Count);
    }

    #endregion

    #region Count Parameter

    [SkippableTheory]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(5, true)]
    [InlineData(10, true)]
    public void Execute_WithCount_AddsCorrectNumberOfPages(int count, bool requiresLicense)
    {
        if (requiresLicense)
            SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", count }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(count.ToString(), result.Message);
        Assert.Equal(initialCount + count, doc.Pages.Count);
        AssertModified(context);
    }

    /// <summary>
    ///     A non-positive count is refused, and the document is left alone.
    ///     <para>
    ///         This case was called "AddsDefaultOnePage" and asserted only
    ///         <c>Pages.Count &gt;= initialCount</c>, which holds for any outcome including the
    ///         real one: the loop ran zero times, nothing was added, the document was still marked
    ///         modified, and the response said "Added 0 page(s)" (R3-C04). The name described
    ///         behaviour that never existed.
    ///     </para>
    /// </summary>
    /// <param name="count">A count that cannot describe pages to add.</param>
    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Execute_WithANonPositiveCount_ShouldBeRefused(int count)
    {
        var doc = CreateEmptyDocument();
        var initialCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", count }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Equal(initialCount, doc.Pages.Count);
    }

    #endregion

    #region InsertAt Parameter

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_WithInsertAt_InsertsAtCorrectPosition(int insertAt)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertAt", insertAt }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Equal(4, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInsertAtBeyondEnd_AppendsToEnd()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertAt", 100 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Equal(4, doc.Pages.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithInsertAtZero_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "insertAt", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Equal(4, doc.Pages.Count);
        AssertModified(context);
    }

    #endregion

    #region Page Size

    [Theory]
    [InlineData(612.0, 792.0)]
    [InlineData(595.0, 842.0)]
    [InlineData(841.0, 1190.0)]
    public void Execute_WithCustomSize_SetsCorrectPageSize(double width, double height)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", width },
            { "height", height }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        AssertModified(context);

        var page = doc.Pages[doc.Pages.Count];
        Assert.Equal(width, page.Rect.Width, 1);
        Assert.Equal(height, page.Rect.Height, 1);
    }

    /// <summary>
    ///     R5-C01: a caller who gives one dimension has it applied, and only the missing one falls
    ///     back to A4. Both were discarded whenever either was absent, and these tests asserted a
    ///     success message rather than the geometry, so the page came out plain A4 and nothing said
    ///     so.
    /// </summary>
    [Fact]
    public void Execute_WithOnlyWidth_KeepsTheWidthAndDefaultsTheHeight()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", 612.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Added", result.Message);
        AssertModified(context);

        var page = doc.Pages[doc.Pages.Count];
        Assert.Equal(612.0, page.Rect.Width, 1);
        Assert.Equal(PageSize.A4.Height, page.Rect.Height, 1);
    }

    [Fact]
    public void Execute_WithOnlyHeight_KeepsTheHeightAndDefaultsTheWidth()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "height", 792.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Added", result.Message);
        AssertModified(context);

        var page = doc.Pages[doc.Pages.Count];
        Assert.Equal(PageSize.A4.Width, page.Rect.Width, 1);
        Assert.Equal(792.0, page.Rect.Height, 1);
    }

    /// <summary>
    ///     A single dimension is validated on its own too. Validation was skipped entirely whenever
    ///     the other was absent, so an unusable width reached the library instead of being refused.
    /// </summary>
    [Theory]
    [InlineData("width")]
    [InlineData("height")]
    public void Execute_WithOneUnusableDimension_ShouldBeRefused(string dimension)
    {
        var doc = CreateEmptyDocument();
        var before = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { dimension, -1.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Equal(before, doc.Pages.Count);
    }

    #endregion

    #region Result Message

    [SkippableTheory]
    [InlineData(1, 1, false)]
    [InlineData(3, 2, true)]
    [InlineData(5, 3, true)]
    public void Execute_ReturnsTotalPageCount(int initialPages, int addedPages, bool requiresLicense)
    {
        if (requiresLicense)
            SkipInEvaluationMode(AsposeLibraryType.Pdf, "Exceeds evaluation mode page limit");
        var doc = CreateDocumentWithPages(initialPages);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", addedPages }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        var expectedTotal = initialPages + addedPages;
        Assert.Contains(expectedTotal.ToString(), result.Message);
    }

    [Fact]
    public void Execute_ReturnsAddedPageCount()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3", result.Message);
        Assert.Contains("Added", result.Message);
    }

    #endregion

    #region Document State Preservation

    [Fact]
    public void Execute_PreservesExistingPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(4, doc.Pages.Count);
    }

    [Fact]
    public void Execute_AddsPageAtEnd_ByDefault()
    {
        var doc = CreateDocumentWithPages(3);
        var initialPageCount = doc.Pages.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialPageCount + 1, doc.Pages.Count);
    }

    #endregion

    #region Page Size Bounds

    /// <summary>
    ///     A page size that is not a finite positive measurement must be refused before it reaches
    ///     <c>SetPageSize</c>. NaN and infinity are not sizes, a non-positive page is not a page,
    ///     and a page far larger than the format allows is a resource request rather than a
    ///     document (R2-R06).
    /// </summary>
    /// <param name="width">Requested width in points.</param>
    /// <param name="height">Requested height in points.</param>
    [Theory]
    [InlineData(double.NaN, 792)]
    [InlineData(612, double.NaN)]
    [InlineData(double.PositiveInfinity, 792)]
    [InlineData(612, double.PositiveInfinity)]
    [InlineData(0, 792)]
    [InlineData(612, 0)]
    [InlineData(-612, 792)]
    [InlineData(1_000_000, 1_000_000)]
    public void Execute_WithUnusablePageSize_ShouldBeRejected(double width, double height)
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", width },
            { "height", height }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithOrdinaryPageSize_ShouldBeAccepted()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", 612.0 },
            { "height", 792.0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, doc.Pages.Count);
    }

    #endregion

    #region Page Geometry

    /// <summary>
    ///     The size was applied after the page had been added, so a refused size left a blank page
    ///     behind while the document was reported unmodified (R4-R05).
    /// </summary>
    /// <param name="width">Requested width in points.</param>
    /// <param name="height">Requested height in points.</param>
    [Theory]
    [InlineData(0d, 100d)]
    [InlineData(100d, 0d)]
    [InlineData(-1d, 100d)]
    [InlineData(double.NaN, 100d)]
    [InlineData(double.PositiveInfinity, 100d)]
    public void Execute_WithAnUnusablePageSize_ShouldLeaveNoPageBehind(double width, double height)
    {
        var document = CreateEmptyDocument();
        var before = document.Pages.Count;
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", 1 },
            { "width", width },
            { "height", height }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Equal(before, document.Pages.Count);
    }

    /// <summary>
    ///     The same on the insert path, which adds at a position rather than at the end.
    /// </summary>
    [Fact]
    public void Execute_WithAnUnusablePageSizeOnInsert_ShouldLeaveNoPageBehind()
    {
        var document = CreateEmptyDocument();
        document.Pages.Add();
        var before = document.Pages.Count;
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "count", 1 },
            { "insertAt", 1 },
            { "width", 0d },
            { "height", 100d }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Equal(before, document.Pages.Count);
    }

    #endregion
}
