using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class InsertFieldWordHandlerTests : WordHandlerTestBase
{
    private readonly InsertFieldWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertField()
    {
        Assert.Equal("insert_field", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithParagraphs(int count)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < count; i++) builder.Writeln($"Paragraph {i + 1}");
        return doc;
    }

    #endregion

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsDateField()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldType", "DATE" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(doc.Range.Fields.Count > 0);
        AssertModified(context);
    }

    [Fact]
    public void Execute_InsertsPageField()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldType", "PAGE" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFieldArgument_InsertsFieldWithArgument()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldType", "DATE" },
            { "fieldArgument", @"\@ ""MMMM d, yyyy""" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("argument", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithParagraphIndex_InsertsAtSpecificParagraph()
    {
        var doc = CreateDocumentWithParagraphs(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldType", "TIME" },
            { "paragraphIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithParagraphIndexMinusOne_InsertsAtEnd()
    {
        var doc = CreateDocumentWithParagraphs(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldType", "NUMPAGES" },
            { "paragraphIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("inserted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFieldType_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldType", "DATE" },
            { "paragraphIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
