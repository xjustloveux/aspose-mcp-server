using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FormField;

public class EditPdfFormFieldHandlerTests : PdfHandlerTestBase
{
    private readonly EditPdfFormFieldHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Edit TextBox

    [Fact]
    public void Execute_WithValue_ChangesTextFieldValue()
    {
        var doc = CreateDocumentWithTextField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "testField" },
            { "value", "Updated Text" }
        });

        _handler.Execute(context, parameters);

        var field = doc.Form["testField"] as TextBoxField;
        Assert.Equal("Updated Text", field?.Value);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsFormField()
    {
        var doc = CreateDocumentWithTextField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "testField" },
            { "value", "New Value" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Edited", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsFieldName()
    {
        var doc = CreateDocumentWithTextField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "testField" },
            { "value", "New Value" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("testField", result.Message);
    }

    #endregion

    #region Edit Checkbox

    [Fact]
    public void Execute_WithCheckedValue_ChangesCheckboxState()
    {
        var doc = CreateDocumentWithCheckbox();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "checkField" },
            { "checkedValue", true }
        });

        _handler.Execute(context, parameters);

        var field = doc.Form["checkField"] as CheckboxField;
        Assert.True(field?.Checked);
    }

    [Fact]
    public void Execute_WithCheckedValueFalse_UnchecksCheckbox()
    {
        var doc = CreateDocumentWithCheckbox();
        var checkbox = doc.Form["checkField"] as CheckboxField;
        checkbox!.Checked = true;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "checkField" },
            { "checkedValue", false }
        });

        _handler.Execute(context, parameters);

        Assert.False(checkbox.Checked);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFieldName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTextField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "value", "New Value" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("fieldName", ex.Message);
    }

    [Fact]
    public void Execute_WithNonExistentFieldName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTextField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "nonExistent" },
            { "value", "New Value" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not found", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTextField()
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        var field = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "testField",
            Value = "Initial Value"
        };
        doc.Form.Add(field);
        return doc;
    }

    private static Document CreateDocumentWithCheckbox()
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        var field = new CheckboxField(page, new Rectangle(100, 700, 120, 720))
        {
            PartialName = "checkField",
            Checked = false
        };
        doc.Form.Add(field);
        return doc;
    }

    #endregion
}
