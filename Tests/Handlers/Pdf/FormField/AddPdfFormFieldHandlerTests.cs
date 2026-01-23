using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Handlers.Pdf.FormField;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.FormField;

public class AddPdfFormFieldHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfFormFieldHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFormField()
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        var field = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "existingField"
        };
        doc.Form.Add(field);
        return doc;
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTextField()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "fieldType", "textbox" },
            { "fieldName", "txtName" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 20.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added", result.Message);
        Assert.Contains("textbox", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsFieldName()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "fieldType", "textbox" },
            { "fieldName", "MyField" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 20.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("MyField", result.Message);
    }

    [Fact]
    public void Execute_AddsFieldToForm()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "fieldType", "textbox" },
            { "fieldName", "testField" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 20.0 }
        });

        _handler.Execute(context, parameters);

        Assert.Single(doc.Form);
    }

    #endregion

    #region Different Field Types

    [Fact]
    public void Execute_AddsCheckboxField()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "fieldType", "checkbox" },
            { "fieldName", "chkAgree" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 20.0 },
            { "height", 20.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("checkbox", result.Message);
    }

    [Fact]
    public void Execute_AddsRadioButtonField()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "fieldType", "radiobutton" },
            { "fieldName", "rdoChoice" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 20.0 },
            { "height", 20.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("radiobutton", result.Message);
    }

    [Fact]
    public void Execute_WithDefaultValue_SetsValue()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "fieldType", "textbox" },
            { "fieldName", "txtDefault" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 20.0 },
            { "defaultValue", "Default Text" }
        });

        _handler.Execute(context, parameters);

        var field = doc.Form["txtDefault"] as TextBoxField;
        Assert.Equal("Default Text", field?.Value);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldType", "textbox" },
            { "fieldName", "txtField" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 20.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 },
            { "fieldType", "textbox" },
            { "fieldName", "txtField" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 20.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithDuplicateFieldName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFormField();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "fieldType", "textbox" },
            { "fieldName", "existingField" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 20.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void Execute_WithUnknownFieldType_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "fieldType", "unknowntype" },
            { "fieldName", "txtField" },
            { "x", 100.0 },
            { "y", 700.0 },
            { "width", 200.0 },
            { "height", 20.0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Unknown field type", ex.Message);
    }

    #endregion
}
