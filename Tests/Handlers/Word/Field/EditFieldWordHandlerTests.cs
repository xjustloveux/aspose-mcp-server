using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class EditFieldWordHandlerTests : WordHandlerTestBase
{
    private readonly EditFieldWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditField()
    {
        Assert.Equal("edit_field", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsField()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithLockField_LocksField()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 },
            { "lockField", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("locked", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(doc.Range.Fields[0].IsLocked);
    }

    [Fact]
    public void Execute_WithUnlockField_UnlocksField()
    {
        var doc = CreateDocumentWithLockedField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 },
            { "unlockField", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("unlocked", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(doc.Range.Fields[0].IsLocked);
    }

    [Fact]
    public void Execute_WithUpdateFieldFalse_DoesNotUpdate()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 },
            { "updateField", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited", result.Message, StringComparison.OrdinalIgnoreCase);
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

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithField()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE");
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
