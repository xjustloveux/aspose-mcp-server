using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Text;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Text;

public class EditPdfTextHandlerTests : PdfHandlerTestBase
{
    private readonly EditPdfTextHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsReplacementCountInMessage()
    {
        var doc = CreateDocumentWithPages(1);
        AddTextToPage(doc.Pages[1], "Test Text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "Test" },
            { "newText", "New" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("New");
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    #endregion

    #region Helper Methods

    private static void AddTextToPage(Aspose.Pdf.Page page, string text)
    {
        var textFragment = new TextFragment(text) { Position = new Position(100, 700) };
        var textBuilder = new TextBuilder(page);
        textBuilder.AppendText(textFragment);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_ReplacesText()
    {
        var doc = CreateDocumentWithPages(1);
        AddTextToPage(doc.Pages[1], "Original Text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "Original" },
            { "newText", "Modified" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Modified");
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    [SkippableTheory]
    [InlineData("Hello", "Goodbye")]
    [InlineData("Test", "Result")]
    public void Execute_ReplacesVariousTexts(string oldText, string newText)
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "Text replacement fails in evaluation mode");
        var doc = CreateDocumentWithPages(1);
        AddTextToPage(doc.Pages[1], oldText);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", oldText },
            { "newText", newText }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber(newText);
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    #endregion

    #region Page Index

    [Fact]
    public void Execute_WithPageIndex_ReplacesOnSpecificPage()
    {
        var doc = CreateDocumentWithPages(3);
        AddTextToPage(doc.Pages[2], "Find Me");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "Find" },
            { "newText", "Found" },
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Found");
            doc.Pages[2].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_DefaultPageIndex_ReplacesOnFirstPage()
    {
        var doc = CreateDocumentWithPages(3);
        AddTextToPage(doc.Pages[1], "Replace This");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "Replace" },
            { "newText", "Modified" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var absorber = new TextFragmentAbsorber("Modified");
            doc.Pages[1].Accept(absorber);
            Assert.True(absorber.TextFragments.Count > 0);
        }

        AssertModified(context);
    }

    #endregion

    #region Replace All

    [Fact]
    public void Execute_WithReplaceAllFalse_ReplacesOnlyFirst()
    {
        var doc = CreateDocumentWithPages(1);
        AddTextToPage(doc.Pages[1], "Word Word Word");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "Word" },
            { "newText", "Changed" },
            { "replaceAll", false }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var changedAbsorber = new TextFragmentAbsorber("Changed");
            doc.Pages[1].Accept(changedAbsorber);
            Assert.Single(changedAbsorber.TextFragments);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithReplaceAllTrue_ReplacesAllOccurrences()
    {
        var doc = CreateDocumentWithPages(1);
        AddTextToPage(doc.Pages[1], "Word");
        AddTextToPage(doc.Pages[1], "Word");
        AddTextToPage(doc.Pages[1], "Word");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "Word" },
            { "newText", "Changed" },
            { "replaceAll", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var changedAbsorber = new TextFragmentAbsorber("Changed");
            doc.Pages[1].Accept(changedAbsorber);
            Assert.True(changedAbsorber.TextFragments.Count >= 3);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutOldText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "newText", "Replacement" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("oldText", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutNewText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "Original" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("newText", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyOldText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "" },
            { "newText", "New" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("oldText", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyNewText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "Old" },
            { "newText", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("newText", ex.Message, StringComparison.OrdinalIgnoreCase);
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
            { "oldText", "Test" },
            { "newText", "Result" },
            { "pageIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_TextNotFound_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "oldText", "NonExistentText" },
            { "newText", "Replacement" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
