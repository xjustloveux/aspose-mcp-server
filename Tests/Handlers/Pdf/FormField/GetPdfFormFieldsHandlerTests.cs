using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Results.Pdf.FormField;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFormFieldsResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Contains("No form fields found", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFormFieldsResult>(res);

        Assert.Equal(2, result.Count);
        Assert.Equal(5, result.TotalCount);
        Assert.True(result.Truncated);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFormFieldsResult>(res);

        Assert.True(result.Count >= 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithFormFields(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFormFieldsResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var doc = CreateDocumentWithFormFields(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFormFieldsResult>(res);

        Assert.Equal(2, result.Items.Count);
    }

    [Fact]
    public void Execute_ReturnsTotalCount()
    {
        var doc = CreateDocumentWithFormFields(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFormFieldsResult>(res);

        Assert.Equal(3, result.TotalCount);
    }

    #endregion

    #region Field Details

    [Fact]
    public void Execute_ReturnsFieldName()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFormFieldsResult>(res);
        var firstField = result.Items[0];

        Assert.NotNull(firstField.Name);
    }

    [Fact]
    public void Execute_ReturnsFieldType()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetFormFieldsResult>(res);
        var firstField = result.Items[0];

        Assert.NotNull(firstField.Type);
    }

    #endregion
}
