using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Stamp;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Stamp;

/// <summary>
///     Tests for <see cref="RemovePdfStampHandler" />.
///     Validates stamp removal with various parameters and error handling.
/// </summary>
public class RemovePdfStampHandlerTests : PdfHandlerTestBase
{
    private readonly RemovePdfStampHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Remove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithStampAnnotations(int count)
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        for (var i = 0; i < count; i++)
        {
            var annotation = new StampAnnotation(page, new Rectangle(100 + i * 50, 100, 200 + i * 50, 200))
            {
                Contents = $"Stamp {i + 1}"
            };
            page.Annotations.Add(annotation);
        }

        return doc;
    }

    #endregion

    #region Basic Remove Operations

    [Fact]
    public void Execute_RemovesAllStampsFromPage()
    {
        var doc = CreateDocumentWithStampAnnotations(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Removed all 3 stamp annotation(s)", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_RemovesSpecificStampByIndex()
    {
        var doc = CreateDocumentWithStampAnnotations(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "stampIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("Removed stamp annotation 1", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_OnPageWithNoStamps_ReturnsAppropriateMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("No stamp annotations found", result.Message);
    }

    #endregion

    #region Modification Tracking

    [Fact]
    public void Execute_MarksModifiedWhenStampsRemoved()
    {
        var doc = CreateDocumentWithStampAnnotations(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    [Fact]
    public void Execute_DoesNotMarkModifiedWhenNoStampsFound()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Error Handling

    [Theory]
    [InlineData(-1)]
    [InlineData(0)]
    [InlineData(10)]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException(int invalidPageIndex)
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", invalidPageIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(10)]
    public void Execute_WithInvalidStampIndex_ThrowsArgumentException(int invalidStampIndex)
    {
        var doc = CreateDocumentWithStampAnnotations(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "stampIndex", invalidStampIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
