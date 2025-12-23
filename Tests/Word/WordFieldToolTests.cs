using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordFieldToolTests : WordTestBase
{
    private readonly WordFieldTool _tool = new();

    [Fact]
    public async Task InsertField_ShouldInsertField()
    {
        // Arrange
        var docPath = CreateWordDocument("test_insert_field.docx");
        var outputPath = CreateTestFilePath("test_insert_field_output.docx");
        var arguments = CreateArguments("insert_field", docPath, outputPath);
        arguments["fieldType"] = "Date";
        arguments["fieldCode"] = "DATE";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var fields = doc.Range.Fields;
        Assert.True(fields.Count > 0, "Document should contain at least one field");
    }

    [Fact]
    public async Task GetFields_ShouldReturnAllFields()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var arguments = CreateArguments("get_fields", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Field", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task UpdateField_ShouldUpdateField()
    {
        // Arrange
        var docPath = CreateWordDocument("test_update_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_field_output.docx");
        var arguments = CreateArguments("update_field", docPath, outputPath);
        arguments["fieldIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public async Task UpdateAllFields_ShouldUpdateAllFields()
    {
        // Arrange
        var docPath = CreateWordDocument("test_update_all_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        builder.InsertField("TIME", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_all_fields_output.docx");
        var arguments = CreateArguments("update_all", docPath, outputPath);

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public async Task EditField_ShouldEditField()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_field_output.docx");
        var arguments = CreateArguments("edit_field", docPath, outputPath);
        arguments["fieldIndex"] = 0;
        arguments["fieldCode"] = "DATE \\@ \"yyyy-MM-dd\"";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public async Task DeleteField_ShouldDeleteField()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var fieldsBefore = doc.Range.Fields.Count;
        Assert.True(fieldsBefore > 0, "Field should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_field_output.docx");
        var arguments = CreateArguments("delete_field", docPath, outputPath);
        arguments["fieldIndex"] = 0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var fieldsAfter = resultDoc.Range.Fields.Count;
        Assert.True(fieldsAfter < fieldsBefore,
            $"Field should be deleted. Before: {fieldsBefore}, After: {fieldsAfter}");
    }

    [Fact]
    public async Task GetFieldDetail_ShouldReturnFieldDetail()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_field_detail.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var arguments = CreateArguments("get_field_detail", docPath);
        arguments["fieldIndex"] = 0;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Field", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AddFormField_ShouldAddFormField()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_form_field.docx");
        var outputPath = CreateTestFilePath("test_add_form_field_output.docx");
        var arguments = CreateArguments("add_form_field", docPath, outputPath);
        arguments["formFieldType"] = "TextInput";
        arguments["fieldName"] = "Name";
        arguments["defaultValue"] = "Default";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var formFields = doc.Range.FormFields;
        Assert.True(formFields.Count > 0, "Document should contain at least one form field");
    }

    [Fact]
    public async Task EditFormField_ShouldEditFormField()
    {
        // Arrange
        var docPath = CreateWordDocument("test_edit_form_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_form_field_output.docx");
        var arguments = CreateArguments("edit_form_field", docPath, outputPath);
        arguments["fieldName"] = "Name";
        arguments["value"] = "Updated Value";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var formField = resultDoc.Range.FormFields["Name"];
        Assert.NotNull(formField);
    }

    [Fact]
    public async Task DeleteFormField_ShouldDeleteFormField()
    {
        // Arrange
        var docPath = CreateWordDocument("test_delete_form_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var formFieldsBefore = doc.Range.FormFields.Count;
        Assert.True(formFieldsBefore > 0, "Form field should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_form_field_output.docx");
        var arguments = CreateArguments("delete_form_field", docPath, outputPath);
        arguments["fieldName"] = "Name";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDoc = new Document(outputPath);
        var formField = resultDoc.Range.FormFields["Name"];
        Assert.Null(formField);
    }

    [Fact]
    public async Task GetFormFields_ShouldReturnAllFormFields()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_form_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var arguments = CreateArguments("get_form_fields", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Form", result, StringComparison.OrdinalIgnoreCase);
    }
}