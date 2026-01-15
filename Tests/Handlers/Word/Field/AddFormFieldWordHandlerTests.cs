using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class AddFormFieldWordHandlerTests : WordHandlerTestBase
{
    private static readonly string[] DropDownOptions = ["Option1", "Option2", "Option3"];

    private readonly AddFormFieldWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddFormField()
    {
        Assert.Equal("add_form_field", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsTextInputField()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formFieldType", "TextInput" },
            { "fieldName", "TestField" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.True(doc.Range.FormFields.Count > 0);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithDefaultValue_SetsDefaultValue()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formFieldType", "TextInput" },
            { "fieldName", "TestField" },
            { "defaultValue", "Default Text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    [Fact]
    public void Execute_AddsCheckBoxField()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formFieldType", "CheckBox" },
            { "fieldName", "CheckField" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Contains("checkbox", result.ToLower());
    }

    [Fact]
    public void Execute_WithCheckedValue_SetsCheckedState()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formFieldType", "CheckBox" },
            { "fieldName", "CheckField" },
            { "checkedValue", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
    }

    [Fact]
    public void Execute_AddsDropDownField()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formFieldType", "DropDown" },
            { "fieldName", "DropField" },
            { "options", DropDownOptions }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Contains("dropdown", result.ToLower());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFormFieldType_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "TestField" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutFieldName_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formFieldType", "TextInput" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidFormFieldType_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formFieldType", "Invalid" },
            { "fieldName", "TestField" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithDropDownWithoutOptions_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formFieldType", "DropDown" },
            { "fieldName", "DropField" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
