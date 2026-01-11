using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class GetFormFieldsWordHandlerTests : WordHandlerTestBase
{
    private readonly GetFormFieldsWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetFormFields()
    {
        Assert.Equal("get_form_fields", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
        var doc = CreateDocumentWithTextInput();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithEmptyDocument_ReturnsEmptyFormFields()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_WithTextInputField_ReturnsTextInputInfo()
    {
        var doc = CreateDocumentWithTextInput();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public void Execute_WithCheckBoxField_ReturnsCheckBoxInfo()
    {
        var doc = CreateDocumentWithCheckBox();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
    }

    [Fact]
    public void Execute_WithDropDownField_ReturnsDropDownInfo()
    {
        var doc = CreateDocumentWithDropDown();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTextInput()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("TextInput1", TextFormFieldType.Regular, "", "Default", 0);
        return doc;
    }

    private static Document CreateDocumentWithCheckBox()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertCheckBox("CheckBox1", false, 0);
        return doc;
    }

    private static Document CreateDocumentWithDropDown()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertComboBox("DropDown1", ["Option1", "Option2", "Option3"], 0);
        return doc;
    }

    #endregion
}
