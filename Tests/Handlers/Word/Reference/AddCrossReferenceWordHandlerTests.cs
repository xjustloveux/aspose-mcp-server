using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Handlers.Word.Reference;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Reference;

public class AddCrossReferenceWordHandlerTests : WordHandlerTestBase
{
    private readonly AddCrossReferenceWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddCrossReference()
    {
        Assert.Equal("add_cross_reference", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithBookmark()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Bookmarked content");
        builder.EndBookmark("MyBookmark");
        return doc;
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsCrossReference()
    {
        var doc = CreateDocumentWithBookmark();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "Bookmark" },
            { "targetName", "MyBookmark" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var refFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldRef).ToList();
        Assert.NotEmpty(refFields);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Contains("MyBookmark", refFields[0].GetFieldCode());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithReferenceText_AddsCrossReferenceWithText()
    {
        var doc = CreateDocumentWithBookmark();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "Heading" },
            { "targetName", "Chapter1" },
            { "referenceText", "See " }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var refFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldRef).ToList();
        Assert.NotEmpty(refFields);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            Assert.Contains("Chapter1", refFields[0].GetFieldCode());
            AssertContainsText(doc, "See ");
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithIncludeAboveBelow_AddsCrossReferenceWithPosition()
    {
        var doc = CreateDocumentWithBookmark();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "Figure" },
            { "targetName", "Figure1" },
            { "includeAboveBelow", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var refFields = doc.Range.Fields.Where(f => f.Type == FieldType.FieldRef).ToList();
        Assert.NotEmpty(refFields);
        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            Assert.Contains("Figure1", refFields[0].GetFieldCode());
            AssertContainsText(doc, "(above)");
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutReferenceType_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetName", "MyBookmark" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutTargetName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "Bookmark" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidReferenceType_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceType", "InvalidType" },
            { "targetName", "MyBookmark" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
