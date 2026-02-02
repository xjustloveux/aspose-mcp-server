using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Handlers.Word.Reference;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var tocFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldTOC).ToList();
        Assert.NotEmpty(tocFields);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var tocFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldTOC).ToList();
        Assert.NotEmpty(tocFields);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) AssertContainsText(doc, "Contents");
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var tocFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldTOC).ToList();
        Assert.NotEmpty(tocFields);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("1-2", tocFields[0].GetFieldCode());
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var tocFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldTOC).ToList();
        Assert.NotEmpty(tocFields);
        AssertModified(context);
    }

    #endregion
}
