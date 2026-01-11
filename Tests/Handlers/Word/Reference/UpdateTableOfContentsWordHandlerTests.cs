using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Reference;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Reference;

public class UpdateTableOfContentsWordHandlerTests : WordHandlerTestBase
{
    private readonly UpdateTableOfContentsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_UpdateTableOfContents()
    {
        Assert.Equal("update_table_of_contents", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidTocIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithToc();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tocIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithToc()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertTableOfContents("\\o \"1-3\"");
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content");
        doc.UpdateFields();
        return doc;
    }

    #endregion

    #region Basic Update Operations

    [Fact]
    public void Execute_WithNoTOC_ReturnsInfoMessage()
    {
        var doc = CreateDocumentWithText("Sample text without TOC.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("no table of contents", result.ToLower());
    }

    [Fact]
    public void Execute_WithTOC_UpdatesTOC()
    {
        var doc = CreateDocumentWithToc();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result.ToLower());
        AssertModified(context);
    }

    #endregion
}
