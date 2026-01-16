using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Reference;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Reference;

public class AddTableOfContentsWordHandlerTests : WordHandlerTestBase
{
    private readonly AddTableOfContentsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddTableOfContents()
    {
        Assert.Equal("add_table_of_contents", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithHeadings()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc) { ParagraphFormat = { StyleIdentifier = StyleIdentifier.Heading1 } };
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content for chapter 1.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content for section 1.1.");
        return doc;
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTableOfContents()
    {
        var doc = CreateDocumentWithHeadings();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("table of contents added", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomTitle_AddsTOCWithTitle()
    {
        var doc = CreateDocumentWithHeadings();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Contents" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("table of contents added", result.ToLower());
    }

    [Fact]
    public void Execute_WithMaxLevel_AddsTOCWithMaxLevel()
    {
        var doc = CreateDocumentWithHeadings();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "maxLevel", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    [Fact]
    public void Execute_AtEndPosition_AddsTOCAtEnd()
    {
        var doc = CreateDocumentWithHeadings();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "position", "end" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    #endregion
}
