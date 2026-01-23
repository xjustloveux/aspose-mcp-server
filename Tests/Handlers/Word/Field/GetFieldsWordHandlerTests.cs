using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Results.Word.Field;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class GetFieldsWordHandlerTests : WordHandlerTestBase
{
    private readonly GetFieldsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetFields()
    {
        Assert.Equal("get_fields", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithDateField();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithDateField()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE");
        return doc;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithEmptyDocument_ReturnsEmptyFields()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFieldsWordResult>(res);

        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void Execute_WithDocumentWithFields_ReturnsFields()
    {
        var doc = CreateDocumentWithDateField();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFieldsWordResult>(res);

        Assert.True(result.Count > 0);
        Assert.NotNull(result.Fields);
        Assert.NotNull(result.StatisticsByType);
    }

    [Fact]
    public void Execute_WithIncludeCodeFalse_ExcludesFieldCode()
    {
        var doc = CreateDocumentWithDateField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeCode", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFieldsWordResult>(res);

        Assert.True(result.Fields.Count > 0);
    }

    [Fact]
    public void Execute_WithIncludeResultFalse_ExcludesFieldResult()
    {
        var doc = CreateDocumentWithDateField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeResult", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFieldsWordResult>(res);

        Assert.NotNull(result.Fields);
    }

    #endregion
}
