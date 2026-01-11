using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Handlers.Word.Field;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Field;

public class DeleteFormFieldWordHandlerTests : WordHandlerTestBase
{
    private readonly DeleteFormFieldWordHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteFormField()
    {
        Assert.Equal("delete_form_field", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesFormFieldByName()
    {
        var doc = CreateDocumentWithFormField("TestField");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "TestField" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Equal(0, doc.Range.FormFields.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFieldNames_DeletesMultipleFields()
    {
        var doc = CreateDocumentWithMultipleFormFields();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldNames", new[] { "Field1", "Field2" } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Contains("2", result);
    }

    [Fact]
    public void Execute_WithoutParameters_DeletesAllFormFields()
    {
        var doc = CreateDocumentWithMultipleFormFields();
        var initialCount = doc.Range.FormFields.Count;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Contains(initialCount.ToString(), result);
    }

    [Fact]
    public void Execute_WithNonExistentField_DeletesZero()
    {
        var doc = CreateDocumentWithFormField("TestField");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "NonExistent" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0", result);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFormField(string fieldName)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", "Default", 0);
        return doc;
    }

    private static Document CreateDocumentWithMultipleFormFields()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Field1", TextFormFieldType.Regular, "", "", 0);
        builder.Writeln();
        builder.InsertTextInput("Field2", TextFormFieldType.Regular, "", "", 0);
        builder.Writeln();
        builder.InsertCheckBox("Field3", false, 0);
        return doc;
    }

    #endregion
}
