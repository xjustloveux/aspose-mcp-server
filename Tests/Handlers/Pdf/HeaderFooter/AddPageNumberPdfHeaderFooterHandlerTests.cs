using AsposeMcpServer.Handlers.Pdf.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.HeaderFooter;

/// <summary>
///     Tests for <see cref="AddPageNumberPdfHeaderFooterHandler" />.
/// </summary>
public class AddPageNumberPdfHeaderFooterHandlerTests : PdfHandlerTestBase
{
    private readonly AddPageNumberPdfHeaderFooterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddPageNumber()
    {
        Assert.Equal("add_page_number", _handler.Operation);
    }

    #endregion

    #region Alignment

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    public void Execute_WithVariousAlignments_Succeeds(string alignment)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "alignment", alignment }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    #endregion

    #region Start Page

    [Fact]
    public void Execute_WithCustomStartPage_Succeeds()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "startPage", 5 }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("3 page(s)", successResult.Message);
        AssertModified(context);
    }

    #endregion

    #region Page Range

    [Fact]
    public void Execute_WithPageRange_AppliesOnlyToSelectedPages()
    {
        var doc = CreateDocumentWithPages(4);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageRange", "2-4" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("3 page(s)", successResult.Message);
        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsPageNumbersToAllPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("3 page(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSinglePage_AddsPageNumber()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("1 page(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithMultiPageDocument_AddsToAllPages()
    {
        var doc = CreateDocumentWithPages(4);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("4 page(s)", successResult.Message);
        AssertModified(context);
    }

    #endregion

    #region Custom Format

    [Fact]
    public void Execute_WithCustomFormatString_Succeeds()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "format", "- {0} -" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("2 page(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithDefaultFormat_Succeeds()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    #endregion

    #region Position

    [Fact]
    public void Execute_DefaultPosition_IsFooter()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("page numbers", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithHeaderPosition_Succeeds()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "position", "header" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFooterPosition_Succeeds()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "position", "footer" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    #endregion

    #region Font Size and Margin

    [Fact]
    public void Execute_WithCustomFontSize_Succeeds()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fontSize", 18.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomMargin_Succeeds()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "margin", 30.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    #endregion
}
