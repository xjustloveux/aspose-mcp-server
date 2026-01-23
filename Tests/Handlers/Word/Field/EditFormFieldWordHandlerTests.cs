using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class EditFormFieldWordHandlerTests : WordHandlerTestBase
{
    private readonly EditFormFieldWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditFormField()
    {
        Assert.Equal("edit_form_field", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsTextInputValue()
    {
        var doc = CreateDocumentWithTextInput("TestField");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "TestField" },
            { "value", "New Value" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("New Value", doc.Range.FormFields["TestField"].Result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_EditsCheckBoxValue()
    {
        var doc = CreateDocumentWithCheckBox("CheckField");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "CheckField" },
            { "checkedValue", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(doc.Range.FormFields["CheckField"].Checked);
    }

    [Fact]
    public void Execute_EditsDropDownSelectedIndex()
    {
        var doc = CreateDocumentWithDropDown("DropField");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "DropField" },
            { "selectedIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(1, doc.Range.FormFields["DropField"].DropDownSelectedIndex);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutFieldName_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTextInput("TestField");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "value", "New Value" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentField_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTextInput("TestField");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "NonExistent" },
            { "value", "New Value" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTextInput(string fieldName)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "Default", 0);
        return doc;
    }

    private static Document CreateDocumentWithCheckBox(string fieldName)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertCheckBox(fieldName, false, 0);
        return doc;
    }

    private static Document CreateDocumentWithDropDown(string fieldName)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertComboBox(fieldName, ["Option1", "Option2", "Option3"], 0);
        return doc;
    }

    #endregion
}
