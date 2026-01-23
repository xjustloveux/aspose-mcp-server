using AsposeMcpServer.Handlers.Word.List;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.List;

public class AddWordListItemHandlerTests : WordHandlerTestBase
{
    private readonly AddWordListItemHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddItem()
    {
        Assert.Equal("add_item", _handler.Operation);
    }

    #endregion

    #region List Level Parameter

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithListLevel_ReturnsLevelInfo(int level)
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Item" },
            { "styleName", "List Paragraph" },
            { "listLevel", level }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains($"Level: {level}", result.Message);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsListItem()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New Item" },
            { "styleName", "List Paragraph" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("List item added successfully", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsStyleName()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Item" },
            { "styleName", "List Paragraph" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Style: List Paragraph", result.Message);
    }

    [SkippableFact]
    public void Execute_AddsTextToDocument()
    {
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode adds watermark to text");
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "My List Item" },
            { "styleName", "List Paragraph" }
        });

        _handler.Execute(context, parameters);

        AssertContainsText(doc, "My List Item");
    }

    #endregion

    #region Apply Style Indent Parameter

    [Fact]
    public void Execute_WithApplyStyleIndentTrue_UsesStyleIndent()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Item" },
            { "styleName", "List Paragraph" },
            { "applyStyleIndent", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Using style-defined indent", result.Message);
    }

    [Fact]
    public void Execute_WithApplyStyleIndentFalseAndLevel_UsesManualIndent()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Item" },
            { "styleName", "List Paragraph" },
            { "listLevel", 2 },
            { "applyStyleIndent", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Manually set", result.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "List Paragraph" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutStyleName_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Item" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyText_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "" },
            { "styleName", "List Paragraph" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyStyleName_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Item" },
            { "styleName", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("styleName", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidStyleName_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Item" },
            { "styleName", "NonExistentStyle12345" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    #endregion
}
