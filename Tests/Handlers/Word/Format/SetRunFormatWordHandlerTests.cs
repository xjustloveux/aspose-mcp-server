using AsposeMcpServer.Handlers.Word.Format;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Format;

public class SetRunFormatWordHandlerTests : WordHandlerTestBase
{
    private readonly SetRunFormatWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetRunFormat()
    {
        Assert.Equal("set_run_format", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidRunIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "runIndex", 99 },
            { "bold", true }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsRunFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "bold", true },
            { "italic", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("run format updated", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsSpecificRunFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "runIndex", 0 },
            { "fontSize", 14.0 },
            { "fontName", "Arial" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result.ToLower());
    }

    [Fact]
    public void Execute_SetsAutoColor()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "color", "auto" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("auto", result.ToLower());
    }

    [Fact]
    public void Execute_SetsUnderline()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "underline", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result.ToLower());
    }

    #endregion
}
