using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.HeaderFooter;

/// <summary>
///     Tests for <see cref="RemovePdfHeaderFooterHandler" />.
/// </summary>
public class RemovePdfHeaderFooterHandlerTests : PdfHandlerTestBase
{
    private readonly RemovePdfHeaderFooterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Remove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithStampAnnotations(int pageCount, int stampsPerPage)
    {
        var doc = new Document();
        for (var i = 0; i < pageCount; i++)
        {
            var page = doc.Pages.Add();
            for (var j = 0; j < stampsPerPage; j++)
            {
                var stamp = new StampAnnotation(page, new Rectangle(10 + j * 50, 10, 50 + j * 50, 50));
                page.Annotations.Add(stamp);
            }
        }

        return doc;
    }

    #endregion

    #region Non-Stamp Annotations Preserved

    [Fact]
    public void Execute_PreservesNonStampAnnotations()
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        var textAnnotation = new TextAnnotation(page, new Rectangle(100, 100, 200, 130))
        {
            Title = "Test",
            Contents = "Note"
        };
        page.Annotations.Add(textAnnotation);
        var stampAnnotation = new StampAnnotation(page, new Rectangle(10, 10, 50, 50));
        page.Annotations.Add(stampAnnotation);

        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("1 stamp(s)", successResult.Message);
        Assert.Single(page.Annotations);
        Assert.IsType<TextAnnotation>(page.Annotations[1]);
        AssertModified(context);
    }

    #endregion

    #region Basic Remove Operations

    [Fact]
    public void Execute_RemovesStampAnnotationsFromAllPages()
    {
        var doc = CreateDocumentWithStampAnnotations(2, 1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("2 stamp(s)", successResult.Message);
        Assert.Contains("2 page(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMultipleStampsPerPage_RemovesAll()
    {
        var doc = CreateDocumentWithStampAnnotations(1, 3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("3 stamp(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoStamps_ReturnsZeroRemoved()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("0 stamp(s)", successResult.Message);
        AssertNotModified(context);
    }

    #endregion

    #region Page Range

    [Fact]
    public void Execute_WithSpecificPageRange_RemovesOnlyFromSelectedPages()
    {
        var doc = CreateDocumentWithStampAnnotations(3, 1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageRange", "1-2" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("2 stamp(s)", successResult.Message);
        Assert.Contains("2 page(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSinglePageRange_RemovesOnlyFromThatPage()
    {
        var doc = CreateDocumentWithStampAnnotations(3, 2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageRange", "2" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("2 stamp(s)", successResult.Message);
        Assert.Contains("1 page(s)", successResult.Message);
        AssertModified(context);
    }

    #endregion
}
