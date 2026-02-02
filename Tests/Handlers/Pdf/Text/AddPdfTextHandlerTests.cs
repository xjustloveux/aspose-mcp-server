using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Text;

public class AddPdfTextHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Position

    [Fact]
    public void Execute_WithPosition_AddsAtSpecifiedCoordinates()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Positioned Text" },
            { "x", 200.0 },
            { "y", 500.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Positioned Text");
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsPageIndexInMessage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test Text" },
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Test Text");
            doc.Pages[2].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTextToPage()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Hello World" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Hello World");
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    [Theory]
    [InlineData("Simple text")]
    [InlineData("Text with numbers 12345")]
    [InlineData("Special chars: @#$%")]
    public void Execute_AddsVariousTexts(string text)
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", text }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber(text);
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    #endregion

    #region Page Index

    [Fact]
    public void Execute_WithPageIndex_AddsToSpecificPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Page 2 Text" },
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Page 2 Text");
            doc.Pages[2].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Execute_AddsToVariousPageIndices(int pageIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" },
            { "pageIndex", pageIndex }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Test");
            doc.Pages[pageIndex].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_DefaultPageIndex_AddsToFirstPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "First Page" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("First Page");
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    #endregion

    #region Font Settings

    [Fact]
    public void Execute_WithFontSettings_AppliesFont()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Styled Text" },
            { "fontName", "Times New Roman" },
            { "fontSize", 16.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Styled Text");
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithColor_AppliesColor()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Colored Text" },
            { "color", "Red" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Colored Text");
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Test" },
            { "pageIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
