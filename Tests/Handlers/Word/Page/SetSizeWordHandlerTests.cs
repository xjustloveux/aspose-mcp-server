using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Page;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Page;

public class SetSizeWordHandlerTests : WordHandlerTestBase
{
    private readonly SetSizeWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetSize()
    {
        Assert.Equal("set_size", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsPaperSizeA4()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paperSize", "A4" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("size updated", result.ToLower());
        Assert.Equal(PaperSize.A4, doc.Sections[0].PageSetup.PaperSize);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsPaperSizeLetter()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paperSize", "Letter" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(PaperSize.Letter, doc.Sections[0].PageSetup.PaperSize);
    }

    [Fact]
    public void Execute_SetsCustomDimensions()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", 612.0 },
            { "height", 792.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("size updated", result.ToLower());
        Assert.Equal(612.0, doc.Sections[0].PageSetup.PageWidth);
        Assert.Equal(792.0, doc.Sections[0].PageSetup.PageHeight);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNoPaperSizeOrDimensions_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithOnlyWidth_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "width", 612.0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
