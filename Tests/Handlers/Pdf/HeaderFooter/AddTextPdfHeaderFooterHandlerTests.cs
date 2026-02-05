using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.HeaderFooter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.HeaderFooter;

/// <summary>
///     Tests for <see cref="AddTextPdfHeaderFooterHandler" />.
/// </summary>
public class AddTextPdfHeaderFooterHandlerTests : PdfHandlerTestBase
{
    private readonly AddTextPdfHeaderFooterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddText()
    {
        Assert.Equal("add_text", _handler.Operation);
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
            { "text", "Aligned Text" },
            { "alignment", alignment }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    #endregion

    #region ResolveHorizontalAlignment

    [Theory]
    [InlineData("left", HorizontalAlignment.Left)]
    [InlineData("right", HorizontalAlignment.Right)]
    [InlineData("center", HorizontalAlignment.Center)]
    [InlineData("unknown", HorizontalAlignment.Center)]
    public void ResolveHorizontalAlignment_ReturnsExpected(string input, HorizontalAlignment expected)
    {
        var result = AddTextPdfHeaderFooterHandler.ResolveHorizontalAlignment(input);
        Assert.Equal(expected, result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTextStampToAllPagesByDefault()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Header Text" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("3 page(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSinglePage_AddsStamp()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test Header" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("1 page(s)", successResult.Message);
        AssertModified(context);
    }

    #endregion

    #region Position

    [Fact]
    public void Execute_WithHeaderPosition_ReturnsHeaderMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Header" },
            { "position", "header" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("header", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFooterPosition_ReturnsFooterMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Footer" },
            { "position", "footer" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("footer", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DefaultPosition_IsHeader()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Default Position" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("header", successResult.Message);
    }

    #endregion

    #region Page Range

    [Fact]
    public void Execute_WithSpecificPageRange_AppliesOnlyToSelectedPages()
    {
        var doc = CreateDocumentWithPages(4);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Range Header" },
            { "pageRange", "1-3" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("3 page(s)", successResult.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCommaPageRange_AppliesOnlyToSelectedPages()
    {
        var doc = CreateDocumentWithPages(4);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Comma Range" },
            { "pageRange", "1,3" }
        });

        var result = _handler.Execute(context, parameters);

        var successResult = Assert.IsType<SuccessResult>(result);
        Assert.Contains("2 page(s)", successResult.Message);
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
            { "text", "Large Header" },
            { "fontSize", 24.0 }
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
            { "text", "Margin Header" },
            { "margin", 40.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(result);
        AssertModified(context);
    }

    #endregion
}
