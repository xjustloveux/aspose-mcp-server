using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FormField;

public class DeletePdfFormFieldHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfFormFieldHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
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

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesFormField()
    {
        var doc = CreateDocumentWithFormFields(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "field0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsFieldName()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "field0" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("field0", result);
    }

    [Fact]
    public void Execute_ReducesFormFieldCount()
    {
        var doc = CreateDocumentWithFormFields(3);
        var initialCount = doc.Form.Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "field0" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount - 1, doc.Form.Count);
    }

    [Fact]
    public void Execute_RemovesCorrectField()
    {
        var doc = CreateDocumentWithFormFields(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "field0" }
        });

        _handler.Execute(context, parameters);

        Assert.Throws<ArgumentException>(() => doc.Form["field0"]);
        Assert.NotNull(doc.Form["field1"]);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFieldName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("fieldName", ex.Message);
    }

    [Fact]
    public void Execute_WithNonExistentFieldName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFormFields(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "nonExistent" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Execute_NoFormFields_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "anyField" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    #endregion
}
