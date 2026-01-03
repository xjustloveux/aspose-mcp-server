using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

public class WordFieldToolTests : WordTestBase
{
    private readonly WordFieldTool _tool;

    public WordFieldToolTests()
    {
        _tool = new WordFieldTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void InsertField_ShouldInsertField()
    {
        var docPath = CreateWordDocument("test_insert_field.docx");
        var outputPath = CreateTestFilePath("test_insert_field_output.docx");
        _tool.Execute("insert_field", docPath, outputPath: outputPath, fieldType: "DATE");
        var doc = new Document(outputPath);
        var fields = doc.Range.Fields;
        Assert.True(fields.Count > 0, "Document should contain at least one field");
    }

    [Fact]
    public void GetFields_ShouldReturnAllFields()
    {
        var docPath = CreateWordDocument("test_get_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);
        var result = _tool.Execute("get_fields", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Field", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UpdateField_ShouldUpdateField()
    {
        var docPath = CreateWordDocument("test_update_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_field_output.docx");
        _tool.Execute("update_field", docPath, outputPath: outputPath, fieldIndex: 0);
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public void UpdateAllFields_ShouldUpdateAllFields()
    {
        var docPath = CreateWordDocument("test_update_all_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        builder.InsertField("TIME", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_all_fields_output.docx");
        _tool.Execute("update_all", docPath, outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public void EditField_ShouldEditField()
    {
        var docPath = CreateWordDocument("test_edit_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_field_output.docx");
        _tool.Execute("edit_field", docPath, outputPath: outputPath,
            fieldIndex: 0, fieldCode: "DATE \\@ \"yyyy-MM-dd\"");
        Assert.True(File.Exists(outputPath), "Output document should be created");
    }

    [Fact]
    public void DeleteField_ShouldDeleteField()
    {
        var docPath = CreateWordDocument("test_delete_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var fieldsBefore = doc.Range.Fields.Count;
        Assert.True(fieldsBefore > 0, "Field should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_field_output.docx");
        _tool.Execute("delete_field", docPath, outputPath: outputPath, fieldIndex: 0);
        var resultDoc = new Document(outputPath);
        var fieldsAfter = resultDoc.Range.Fields.Count;
        Assert.True(fieldsAfter < fieldsBefore,
            $"Field should be deleted. Before: {fieldsBefore}, After: {fieldsAfter}");
    }

    [Fact]
    public void GetFieldDetail_ShouldReturnFieldDetail()
    {
        var docPath = CreateWordDocument("test_get_field_detail.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);
        var result = _tool.Execute("get_field_detail", docPath, fieldIndex: 0);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Field", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddFormField_ShouldAddFormField()
    {
        var docPath = CreateWordDocument("test_add_form_field.docx");
        var outputPath = CreateTestFilePath("test_add_form_field_output.docx");
        _tool.Execute("add_form_field", docPath, outputPath: outputPath,
            formFieldType: "TextInput", fieldName: "Name", defaultValue: "Default");
        var doc = new Document(outputPath);
        var formFields = doc.Range.FormFields;
        Assert.True(formFields.Count > 0, "Document should contain at least one form field");
        Assert.NotNull(formFields["Name"]);
    }

    [Fact]
    public void EditFormField_ShouldEditFormField()
    {
        var docPath = CreateWordDocument("test_edit_form_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_form_field_output.docx");
        _tool.Execute("edit_form_field", docPath, outputPath: outputPath,
            fieldName: "Name", value: "Updated Value");
        var resultDoc = new Document(outputPath);
        var formField = resultDoc.Range.FormFields["Name"];
        Assert.NotNull(formField);
    }

    [Fact]
    public void DeleteFormField_ShouldDeleteFormField()
    {
        var docPath = CreateWordDocument("test_delete_form_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var formFieldsBefore = doc.Range.FormFields.Count;
        Assert.True(formFieldsBefore > 0, "Form field should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_form_field_output.docx");
        _tool.Execute("delete_form_field", docPath, outputPath: outputPath, fieldName: "Name");
        var resultDoc = new Document(outputPath);
        var formField = resultDoc.Range.FormFields["Name"];
        Assert.Null(formField);
    }

    [Fact]
    public void GetFormFields_ShouldReturnAllFormFields()
    {
        var docPath = CreateWordDocument("test_get_form_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);
        var result = _tool.Execute("get_form_fields", docPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Form", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_unknown_op.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void InsertField_WithEmptyFieldType_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_insert_empty_type.docx");
        var outputPath = CreateTestFilePath("test_insert_empty_type_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_field", docPath, outputPath: outputPath, fieldType: ""));

        Assert.Contains("fieldType is required", ex.Message);
    }

    [Fact]
    public void UpdateField_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_update_invalid_index.docx");
        var outputPath = CreateTestFilePath("test_update_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("update_field", docPath, outputPath: outputPath, fieldIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteField_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_delete_invalid_index.docx");
        var outputPath = CreateTestFilePath("test_delete_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_field", docPath, outputPath: outputPath, fieldIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditField_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_invalid_index.docx");
        var outputPath = CreateTestFilePath("test_edit_invalid_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_field", docPath, outputPath: outputPath, fieldIndex: 999, fieldCode: "DATE"));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetFieldDetail_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_get_detail_invalid_index.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_field_detail", docPath, fieldIndex: 999));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void AddFormField_WithEmptyFieldName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_form_empty_name.docx");
        var outputPath = CreateTestFilePath("test_add_form_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_form_field", docPath, outputPath: outputPath,
                formFieldType: "TextInput", fieldName: ""));

        Assert.Contains("fieldName is required", ex.Message);
    }

    [Fact]
    public void EditFormField_WithNonExistentFieldName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_form_nonexistent.docx");
        var outputPath = CreateTestFilePath("test_edit_form_nonexistent_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_form_field", docPath, outputPath: outputPath,
                fieldName: "NonExistent", value: "Value"));

        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void DeleteFormField_WithNonExistentFieldName_ShouldDeleteZeroFields()
    {
        var docPath = CreateWordDocument("test_delete_form_nonexistent.docx");
        var outputPath = CreateTestFilePath("test_delete_form_nonexistent_output.docx");

        // Act - Non-existent field name results in 0 deletions (no exception)
        var result = _tool.Execute("delete_form_field", docPath, outputPath: outputPath, fieldName: "NonExistent");
        Assert.True(File.Exists(outputPath));
        Assert.Contains("Deleted 0 form field(s)", result);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void InsertField_WithSessionId_ShouldInsertFieldInMemory()
    {
        var docPath = CreateWordDocument("test_session_insert_field.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("insert_field", sessionId: sessionId, fieldType: "DATE");
        Assert.Contains("field", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory document has the field
        var doc = SessionManager.GetDocument<Document>(sessionId);
        var fields = doc.Range.Fields;
        Assert.True(fields.Count > 0, "Session document should contain at least one field");
    }

    [Fact]
    public void GetFields_WithSessionId_ShouldReturnFields()
    {
        var docPath = CreateWordDocument("test_session_get_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_fields", sessionId: sessionId);
        Assert.Contains("Field", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteField_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);

        // Verify field exists before deletion
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var fieldsBefore = sessionDoc.Range.Fields.Count;
        Assert.True(fieldsBefore > 0, "Field should exist before deletion");
        _tool.Execute("delete_field", sessionId: sessionId, fieldIndex: 0);

        // Assert - verify in-memory deletion
        var fieldsAfter = sessionDoc.Range.Fields.Count;
        Assert.True(fieldsAfter < fieldsBefore, "Field should be deleted in session");
    }

    [Fact]
    public void AddFormField_WithSessionId_ShouldAddFormFieldInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_form_field.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_form_field", sessionId: sessionId,
            formFieldType: "TextInput", fieldName: "SessionField", defaultValue: "Default");

        // Assert - Format is "{formFieldType} field '{fieldName}' added"
        Assert.Contains("TextInput field 'SessionField' added", result);

        // Verify in-memory document has the form field
        var doc = SessionManager.GetDocument<Document>(sessionId);
        var formFields = doc.Range.FormFields;
        Assert.True(formFields.Count > 0, "Session document should contain at least one form field");
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_fields", sessionId: "invalid_session_id"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_path_field.docx");
        var doc1 = new Document(docPath1);
        var builder1 = new DocumentBuilder(doc1);
        builder1.InsertField("AUTHOR", "PathFieldAuthor");
        doc1.Save(docPath1);

        var docPath2 = CreateWordDocument("test_session_field.docx");
        var doc2 = new Document(docPath2);
        var builder2 = new DocumentBuilder(doc2);
        builder2.InsertField("TITLE", "SessionFieldTitle");
        doc2.Save(docPath2);

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId
        var result = _tool.Execute("get_fields", docPath1, sessionId);

        // Assert - should use sessionId, returning TITLE field not AUTHOR field
        Assert.Contains("TITLE", result);
        Assert.DoesNotContain("AUTHOR", result);
    }

    #endregion
}