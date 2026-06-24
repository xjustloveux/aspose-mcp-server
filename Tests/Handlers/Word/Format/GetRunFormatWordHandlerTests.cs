using Aspose.Words;
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
        Assert.Equal("get", _handler.Operation);
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

    #region Story-Relative Address

    [Fact]
    public void Execute_AllRuns_WithStoryTypeHeader_ReportsHeaderAddress()
    {
        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        builder.Write("body");
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("header-text");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "storyType", "Header" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRunFormatAllResult>(res);

        Assert.Equal("Header", result.StoryType);
        Assert.Equal("Primary", result.HeaderFooterType);
        Assert.Equal(0, result.ParagraphIndex);
    }

    [Fact]
    public void Execute_SingleRun_WithStoryTypeHeader_ReportsHeaderAddress()
    {
        var doc = CreateEmptyDocument();
        var builder = new DocumentBuilder(doc);
        builder.Write("body");
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Write("header-text");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", 0 },
            { "storyType", "Header" },
            { "runIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRunFormatWordResult>(res);

        Assert.Equal("Header", result.StoryType);
        Assert.Equal("Primary", result.HeaderFooterType);
        Assert.Equal(0, result.ParagraphIndex);
    }

    [Fact]
    public void Execute_AllRuns_WithMinusOneIndex_ReportsResolvedIndexNotRawInput()
    {
        var doc = CreateDocumentWithParagraphs("first", "second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "paragraphIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetRunFormatAllResult>(res);

        Assert.Equal(1, result.ParagraphIndex);
        Assert.Equal("Body", result.StoryType);
        Assert.Null(result.HeaderFooterType);
    }

    #endregion
}
