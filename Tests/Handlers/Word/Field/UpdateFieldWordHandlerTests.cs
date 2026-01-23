using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class UpdateFieldWordHandlerTests : WordHandlerTestBase
{
    private readonly UpdateFieldWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_UpdateField()
    {
        Assert.Equal("update_field", _handler.Operation);
    }

    #endregion

    #region Error Handling

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

    #endregion

    #region Basic Update Operations

    [Fact]
    public void Execute_UpdatesSpecificField()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithUpdateAll_UpdatesAllFields()
    {
        var doc = CreateDocumentWithMultipleFields();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "updateAll", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithLockedField_ReturnsWarning()
    {
        var doc = CreateDocumentWithLockedField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("locked", result.Message, StringComparison.OrdinalIgnoreCase);
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

    private static Document CreateDocumentWithMultipleFields()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE");
        builder.Writeln();
        builder.InsertField("TIME");
        builder.Writeln();
        builder.InsertField("PAGE");
        return doc;
    }

    private static Document CreateDocumentWithLockedField()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var field = builder.InsertField("DATE");
        field.IsLocked = true;
        return doc;
    }

    #endregion
}
