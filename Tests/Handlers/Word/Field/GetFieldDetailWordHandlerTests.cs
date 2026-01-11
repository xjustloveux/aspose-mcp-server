using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class GetFieldDetailWordHandlerTests : WordHandlerTestBase
{
    private readonly GetFieldDetailWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetFieldDetail()
    {
        Assert.Equal("get_field_detail", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 }
        });

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsFieldDetail()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("index", out _));
        Assert.True(json.RootElement.TryGetProperty("type", out _));
        Assert.True(json.RootElement.TryGetProperty("code", out _));
    }

    [Fact]
    public void Execute_ReturnsFieldLockState()
    {
        var doc = CreateDocumentWithField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("isLocked", out _));
        Assert.True(json.RootElement.TryGetProperty("isDirty", out _));
    }

    [Fact]
    public void Execute_WithHyperlinkField_ReturnsHyperlinkDetails()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Contains("hyperlink", json.RootElement.GetProperty("type").GetString()!.ToLower());
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

    private static Document CreateDocumentWithHyperlink()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Click here", "https://example.com", false);
        return doc;
    }

    #endregion
}
