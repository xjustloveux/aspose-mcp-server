using AsposeMcpServer.Handlers.Word.Format;
using AsposeMcpServer.Results.Word.Format;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Format;

public class GetRunFormatWordHandlerTests : WordHandlerTestBase
{
    private readonly GetRunFormatWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetRunFormat()
    {
        Assert.Equal("get_run_format", _handler.Operation);
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
            { "runIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsRunFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRunFormatAllResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.Runs);
        Assert.True(result.Count > 0);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithRunIndex_ReturnsSpecificRunFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "runIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRunFormatWordResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.FontName);
        Assert.True(result.FontSize > 0);
    }

    [Fact]
    public void Execute_WithIncludeInherited_ReturnsInheritedFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "runIndex", 0 },
            { "includeInherited", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRunFormatWordResult>(res);

        Assert.NotNull(result);
        Assert.Equal("inherited", result.FormatType);
    }

    #endregion
}
