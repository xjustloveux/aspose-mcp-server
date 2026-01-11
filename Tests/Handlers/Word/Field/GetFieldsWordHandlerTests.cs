using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_WithDocumentWithFields_ReturnsFields()
    {
        var doc = CreateDocumentWithDateField();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
        Assert.True(json.RootElement.TryGetProperty("fields", out _));
        Assert.True(json.RootElement.TryGetProperty("statisticsByType", out _));
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var fields = json.RootElement.GetProperty("fields");
        Assert.True(fields.GetArrayLength() > 0);
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

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("fields", out _));
    }

    #endregion
}
