using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FormField;

public class GetPdfFormFieldsHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfFormFieldsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region No Form Fields

    [Fact]
    public void Execute_NoFormFields_ReturnsEmptyResult()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No form fields found", json.RootElement.GetProperty("message").GetString());
    }

    #endregion

    #region Limit Parameter

    [Fact]
    public void Execute_WithLimit_LimitsResults()
    {
        var doc = CreateDocumentWithFormFields(5);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "limit", 2 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(5, json.RootElement.GetProperty("totalCount").GetInt32());
        Assert.True(json.RootElement.GetProperty("truncated").GetBoolean());
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFormFields(int count)
    {
        var doc = new Document();
        var page = doc.Pages.Add();

        for (var i = 0; i < count; i++)
        {
            var field = new TextBoxField(page, new Rectangle(100, 700 - i * 30, 300, 720 - i * 30))
            {
                PartialName = $"field{i}"
            };
            doc.Form.Add(field);
        }

        return doc;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsFormFields()
    {
        var doc = CreateDocumentWithFormFields(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("count", out _));
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithFormFields(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var doc = CreateDocumentWithFormFields(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(2, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void Execute_ReturnsTotalCount()
    {
        var doc = CreateDocumentWithFormFields(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(3, json.RootElement.GetProperty("totalCount").GetInt32());
    }

    #endregion

    #region Field Details

    [Fact]
    public void Execute_ReturnsFieldName()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstField = json.RootElement.GetProperty("items")[0];

        Assert.NotNull(firstField.GetProperty("name").GetString());
    }

    [Fact]
    public void Execute_ReturnsFieldType()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstField = json.RootElement.GetProperty("items")[0];

        Assert.NotNull(firstField.GetProperty("type").GetString());
    }

    #endregion
}
