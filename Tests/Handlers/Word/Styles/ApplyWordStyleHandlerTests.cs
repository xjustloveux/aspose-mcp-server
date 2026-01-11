using AsposeMcpServer.Handlers.Word.Styles;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Styles;

public class ApplyWordStyleHandlerTests : WordHandlerTestBase
{
    private readonly ApplyWordStyleHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ApplyStyle()
    {
        Assert.Equal("apply_style", _handler.Operation);
    }

    #endregion

    #region Basic Apply Operations

    [Fact]
    public void Execute_AppliesStyleToParagraph()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Heading 1" },
            { "paragraphIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("applied style", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithApplyToAllParagraphs_AppliesToAll()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Normal" },
            { "applyToAllParagraphs", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("3", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutStyleName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidStyleName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "NonExistentStyle" },
            { "paragraphIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoTargetSpecified_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "Heading 1" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
