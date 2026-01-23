using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class DeleteFieldWordHandlerTests : WordHandlerTestBase
{
    private readonly DeleteFieldWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteField()
    {
        Assert.Equal("delete_field", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithField()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE");
        return doc;
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesField()
    {
        var doc = CreateDocumentWithField();
        var initialCount = doc.Range.Fields.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount - 1, doc.Range.Fields.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithKeepResult_UnlinksField()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 },
            { "keepResult", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("yes", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithoutKeepResult_RemovesField()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 },
            { "keepResult", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("no", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFieldIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidFieldIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeFieldIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
