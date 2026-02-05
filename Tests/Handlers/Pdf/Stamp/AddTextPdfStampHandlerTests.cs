using AsposeMcpServer.Handlers.Pdf.Stamp;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Stamp;

/// <summary>
///     Tests for <see cref="AddTextPdfStampHandler" />.
///     Validates text stamp creation with various parameters and error handling.
/// </summary>
public class AddTextPdfStampHandlerTests : PdfHandlerTestBase
{
    private readonly AddTextPdfStampHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddText()
    {
        Assert.Equal("add_text", _handler.Operation);
    }

    #endregion

    #region Modification Tracking

    [Fact]
    public void Execute_MarksContextAsModified()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "STAMP" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_WithText_AddsStampToAllPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "CONFIDENTIAL" },
            { "pageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("all pages", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithText_AddsStampToSpecificPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "DRAFT" },
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("page 2", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomOpacity_AddsStamp()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "WATERMARK" },
            { "opacity", 0.5 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomRotation_AddsStamp()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "ROTATED" },
            { "rotation", 45.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomFontSize_AddsStamp()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "LARGE TEXT" },
            { "fontSize", 24.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomColor_AddsStamp()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "RED TEXT" },
            { "color", "red" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithPositionXY_AddsStamp()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "POSITIONED" },
            { "x", 100.0 },
            { "y", 200.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAllOptions_AddsStamp()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "FULL OPTIONS" },
            { "pageIndex", 1 },
            { "x", 50.0 },
            { "y", 100.0 },
            { "fontSize", 20.0 },
            { "opacity", 0.7 },
            { "rotation", 30.0 },
            { "color", "blue" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("page 1", result.Message);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException(int invalidPageIndex)
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "TEST" },
            { "pageIndex", invalidPageIndex }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
