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
        Assert.Equal("update", _handler.Operation);
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

    /// <summary>
    ///     R5-C02: an update-all that updated nothing must leave the session clean. The tally was
    ///     read and then the context was marked modified regardless, so a document with no fields —
    ///     or one whose every field was locked or refused — reported "Updated 0 field(s)" and still
    ///     made the close path rewrite an unchanged file.
    /// </summary>
    [Fact]
    public void Execute_WithUpdateAllOnADocumentWithNoFields_ShouldLeaveItUnmodified()
    {
        var context = CreateContext(new Document());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "updateAll", true }
        });

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        Assert.Contains("Updated 0 field(s)", result.Message, StringComparison.Ordinal);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithUpdateAllWhenEveryFieldIsLocked_ShouldLeaveItUnmodified()
    {
        var context = CreateContext(CreateDocumentWithOnlyLockedFields());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "updateAll", true }
        });

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        Assert.Contains("Updated 0 field(s)", result.Message, StringComparison.Ordinal);
        Assert.Contains("locked", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithUpdateAllWhenEveryFieldIsRefused_ShouldLeaveItUnmodified()
    {
        var context = CreateContext(CreateDocumentWithOnlyRefusedFields());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "updateAll", true }
        });

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        Assert.Contains("Updated 0 field(s)", result.Message, StringComparison.Ordinal);
        Assert.Contains("external content", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithOnlyLockedFields()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        foreach (var code in new[] { "DATE", "TIME" })
        {
            builder.InsertField(code).IsLocked = true;
            builder.Writeln();
        }

        return doc;
    }

    private static Document CreateDocumentWithOnlyRefusedFields()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField("INCLUDETEXT \"nowhere.docx\"");
        return doc;
    }

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
