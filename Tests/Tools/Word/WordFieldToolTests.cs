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

    #region General

    [Fact]
    public void InsertField_ShouldInsertField()
    {
        var docPath = CreateWordDocument("test_insert_field.docx");
        var outputPath = CreateTestFilePath("test_insert_field_output.docx");
        _tool.Execute("insert_field", docPath, outputPath: outputPath, fieldType: "DATE");
        var doc = new Document(outputPath);
        var fields = doc.Range.Fields;
        Assert.True(fields.Count > 0);
    }

    [Theory]
    [InlineData("DATE")]
    [InlineData("TIME")]
    [InlineData("PAGE")]
    [InlineData("NUMPAGES")]
    [InlineData("AUTHOR")]
    public void InsertField_WithDifferentFieldTypes_ShouldWork(string fieldType)
    {
        var docPath = CreateWordDocument($"test_insert_{fieldType}.docx");
        var outputPath = CreateTestFilePath($"test_insert_{fieldType}_output.docx");
        var result = _tool.Execute("insert_field", docPath, outputPath: outputPath, fieldType: fieldType);
        Assert.Contains(fieldType, result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void InsertField_WithFieldArgument_ShouldIncludeArgument()
    {
        var docPath = CreateWordDocument("test_insert_with_arg.docx");
        var outputPath = CreateTestFilePath("test_insert_with_arg_output.docx");
        var result = _tool.Execute("insert_field", docPath, outputPath: outputPath,
            fieldType: "DATE", fieldArgument: "\\@ \"yyyy-MM-dd\"");
        Assert.Contains("yyyy-MM-dd", result);
    }

    [Fact]
    public void InsertField_WithParagraphIndex_ShouldInsertAtPosition()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_para.docx", "First paragraph");
        var outputPath = CreateTestFilePath("test_insert_para_output.docx");
        _tool.Execute("insert_field", docPath, outputPath: outputPath,
            fieldType: "DATE", paragraphIndex: 0);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void InsertField_WithInsertAtStart_ShouldInsertAtParagraphStart()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_start.docx", "Content");
        var outputPath = CreateTestFilePath("test_insert_start_output.docx");
        _tool.Execute("insert_field", docPath, outputPath: outputPath,
            fieldType: "DATE", paragraphIndex: 0, insertAtStart: true);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GetFields_ShouldReturnAllFields()
    {
        var docPath = CreateWordDocument("test_get_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        builder.InsertField("TIME", "");
        doc.Save(docPath);
        var result = _tool.Execute("get_fields", docPath);
        Assert.Contains("count", result);
        Assert.Contains("DATE", result);
    }

    [Fact]
    public void GetFields_WithIncludeOptions_ShouldRespectOptions()
    {
        var docPath = CreateWordDocument("test_get_fields_options.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);
        var result = _tool.Execute("get_fields", docPath, includeCode: true, includeResult: false);
        Assert.Contains("code", result);
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
        var result = _tool.Execute("update_field", docPath, outputPath: outputPath, fieldIndex: 0);
        Assert.Contains("updated", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UpdateAllFields_ShouldUpdateAllFields()
    {
        var docPath = CreateWordDocument("test_update_all.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        builder.InsertField("TIME", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_all_output.docx");
        var result = _tool.Execute("update_all", docPath, outputPath: outputPath);
        Assert.StartsWith("Updated", result);
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
        var result = _tool.Execute("edit_field", docPath, outputPath: outputPath,
            fieldIndex: 0, fieldCode: "DATE \\@ \"yyyy-MM-dd\"");
        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EditField_WithLockField_ShouldLockField()
    {
        var docPath = CreateWordDocument("test_edit_lock.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_lock_output.docx");
        var result = _tool.Execute("edit_field", docPath, outputPath: outputPath,
            fieldIndex: 0, lockField: true);
        Assert.Contains("locked", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EditField_WithUnlockField_ShouldUnlockField()
    {
        var docPath = CreateWordDocument("test_edit_unlock.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        var field = builder.InsertField("DATE", "");
        field.IsLocked = true;
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_unlock_output.docx");
        var result = _tool.Execute("edit_field", docPath, outputPath: outputPath,
            fieldIndex: 0, unlockField: true);
        Assert.Contains("unlocked", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteField_ShouldDeleteField()
    {
        var docPath = CreateWordDocument("test_delete_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_field_output.docx");
        _tool.Execute("delete_field", docPath, outputPath: outputPath, fieldIndex: 0);
        var resultDoc = new Document(outputPath);
        Assert.Equal(0, resultDoc.Range.Fields.Count);
    }

    [Fact]
    public void DeleteField_WithKeepResult_ShouldKeepText()
    {
        var docPath = CreateWordDocument("test_delete_keep.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_keep_output.docx");
        var result = _tool.Execute("delete_field", docPath, outputPath: outputPath,
            fieldIndex: 0, keepResult: true);
        Assert.Contains("Keep result text: Yes", result);
    }

    [Fact]
    public void GetFieldDetail_ShouldReturnFieldDetail()
    {
        var docPath = CreateWordDocument("test_get_detail.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);
        var result = _tool.Execute("get_field_detail", docPath, fieldIndex: 0);
        Assert.Contains("type", result);
        Assert.Contains("code", result);
    }

    [Fact]
    public void AddFormField_TextInput_ShouldAddField()
    {
        var docPath = CreateWordDocument("test_add_text.docx");
        var outputPath = CreateTestFilePath("test_add_text_output.docx");
        _tool.Execute("add_form_field", docPath, outputPath: outputPath,
            formFieldType: "TextInput", fieldName: "Name", defaultValue: "Default");
        var doc = new Document(outputPath);
        Assert.NotNull(doc.Range.FormFields["Name"]);
    }

    [Fact]
    public void AddFormField_CheckBox_ShouldAddField()
    {
        var docPath = CreateWordDocument("test_add_checkbox.docx");
        var outputPath = CreateTestFilePath("test_add_checkbox_output.docx");
        _tool.Execute("add_form_field", docPath, outputPath: outputPath,
            formFieldType: "CheckBox", fieldName: "Accept", checkedValue: true);
        var doc = new Document(outputPath);
        var field = doc.Range.FormFields["Accept"];
        Assert.NotNull(field);
        Assert.True(field.Checked);
    }

    [Fact]
    public void AddFormField_DropDown_ShouldAddField()
    {
        var docPath = CreateWordDocument("test_add_dropdown.docx");
        var outputPath = CreateTestFilePath("test_add_dropdown_output.docx");
        _tool.Execute("add_form_field", docPath, outputPath: outputPath,
            formFieldType: "DropDown", fieldName: "Country", options: ["USA", "UK", "Canada"]);
        var doc = new Document(outputPath);
        var field = doc.Range.FormFields["Country"];
        Assert.NotNull(field);
        Assert.Equal(3, field.DropDownItems.Count);
    }

    [Fact]
    public void EditFormField_TextInput_ShouldEditField()
    {
        var docPath = CreateWordDocument("test_edit_text.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_text_output.docx");
        _tool.Execute("edit_form_field", docPath, outputPath: outputPath,
            fieldName: "Name", value: "Updated");
        var resultDoc = new Document(outputPath);
        Assert.NotNull(resultDoc.Range.FormFields["Name"]);
    }

    [Fact]
    public void EditFormField_CheckBox_ShouldEditField()
    {
        var docPath = CreateWordDocument("test_edit_checkbox.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertCheckBox("Accept", false, 0);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_checkbox_output.docx");
        _tool.Execute("edit_form_field", docPath, outputPath: outputPath,
            fieldName: "Accept", checkedValue: true);
        var resultDoc = new Document(outputPath);
        Assert.True(resultDoc.Range.FormFields["Accept"].Checked);
    }

    [Fact]
    public void EditFormField_DropDown_ShouldEditField()
    {
        var docPath = CreateWordDocument("test_edit_dropdown.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertComboBox("Country", ["USA", "UK", "Canada"], 0);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_dropdown_output.docx");
        _tool.Execute("edit_form_field", docPath, outputPath: outputPath,
            fieldName: "Country", selectedIndex: 1);
        var resultDoc = new Document(outputPath);
        Assert.Equal(1, resultDoc.Range.FormFields["Country"].DropDownSelectedIndex);
    }

    [Fact]
    public void DeleteFormField_ShouldDeleteField()
    {
        var docPath = CreateWordDocument("test_delete_form.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_form_output.docx");
        _tool.Execute("delete_form_field", docPath, outputPath: outputPath, fieldName: "Name");
        var resultDoc = new Document(outputPath);
        Assert.Null(resultDoc.Range.FormFields["Name"]);
    }

    [Fact]
    public void DeleteFormField_WithFieldNames_ShouldDeleteMultiple()
    {
        var docPath = CreateWordDocument("test_delete_multi.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        builder.InsertTextInput("Email", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_delete_multi_output.docx");
        var result = _tool.Execute("delete_form_field", docPath, outputPath: outputPath,
            fieldNames: ["Name", "Email"]);
        Assert.StartsWith("Deleted 2", result);
    }

    [Fact]
    public void GetFormFields_ShouldReturnAllFormFields()
    {
        var docPath = CreateWordDocument("test_get_forms.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        builder.InsertCheckBox("Accept", false, 0);
        doc.Save(docPath);
        var result = _tool.Execute("get_form_fields", docPath);
        Assert.Contains("count", result);
        Assert.Contains("Name", result);
        Assert.Contains("Accept", result);
    }

    [Theory]
    [InlineData("INSERT_FIELD")]
    [InlineData("Insert_Field")]
    [InlineData("insert_field")]
    public void Operation_ShouldBeCaseInsensitive_InsertField(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, fieldType: "DATE");
        Assert.StartsWith("Field inserted successfully", result);
    }

    [Theory]
    [InlineData("GET_FIELDS")]
    [InlineData("Get_Fields")]
    [InlineData("get_fields")]
    public void Operation_ShouldBeCaseInsensitive_GetFields(string operation)
    {
        var docPath = CreateWordDocument($"test_case_get_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);
        var result = _tool.Execute(operation, docPath);
        Assert.Contains("count", result);
    }

    [Theory]
    [InlineData("UPDATE_ALL")]
    [InlineData("Update_All")]
    [InlineData("update_all")]
    public void Operation_ShouldBeCaseInsensitive_UpdateAll(string operation)
    {
        var docPath = CreateWordDocument($"test_case_update_{operation.Replace("_", "")}.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);
        var outputPath = CreateTestFilePath($"test_case_update_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath);
        Assert.StartsWith("Updated", result);
    }

    #endregion

    #region Exception

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
        var docPath = CreateWordDocument("test_insert_empty.docx");
        var outputPath = CreateTestFilePath("test_insert_empty_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_field", docPath, outputPath: outputPath, fieldType: ""));
        Assert.Contains("fieldType is required", ex.Message);
    }

    [Fact]
    public void InsertField_WithInvalidParagraphIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocumentWithContent("test_insert_invalid_para.docx", "Content");
        var outputPath = CreateTestFilePath("test_insert_invalid_para_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("insert_field", docPath, outputPath: outputPath,
                fieldType: "DATE", paragraphIndex: 999));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void UpdateField_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_update_invalid.docx");
        var outputPath = CreateTestFilePath("test_update_invalid_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("update_field", docPath, outputPath: outputPath, fieldIndex: 999));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteField_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_delete_invalid.docx");
        var outputPath = CreateTestFilePath("test_delete_invalid_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_field", docPath, outputPath: outputPath, fieldIndex: 999));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditField_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_invalid.docx");
        var outputPath = CreateTestFilePath("test_edit_invalid_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_field", docPath, outputPath: outputPath,
                fieldIndex: 999, fieldCode: "DATE"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditField_WithoutFieldIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_no_index.docx");
        var outputPath = CreateTestFilePath("test_edit_no_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_field", docPath, outputPath: outputPath, fieldCode: "DATE"));
        Assert.Contains("fieldIndex is required", ex.Message);
    }

    [Fact]
    public void DeleteField_WithoutFieldIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_delete_no_index.docx");
        var outputPath = CreateTestFilePath("test_delete_no_index_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_field", docPath, outputPath: outputPath));
        Assert.Contains("fieldIndex is required", ex.Message);
    }

    [Fact]
    public void GetFieldDetail_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_detail_invalid.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_field_detail", docPath, fieldIndex: 999));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetFieldDetail_WithoutFieldIndex_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_detail_no_index.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_field_detail", docPath));
        Assert.Contains("fieldIndex is required", ex.Message);
    }

    [Fact]
    public void AddFormField_WithEmptyFieldName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_empty_name.docx");
        var outputPath = CreateTestFilePath("test_add_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_form_field", docPath, outputPath: outputPath,
                formFieldType: "TextInput", fieldName: ""));
        Assert.Contains("fieldName is required", ex.Message);
    }

    [Fact]
    public void AddFormField_WithEmptyFormFieldType_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_empty_type.docx");
        var outputPath = CreateTestFilePath("test_add_empty_type_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_form_field", docPath, outputPath: outputPath,
                formFieldType: "", fieldName: "Name"));
        Assert.Contains("formFieldType is required", ex.Message);
    }

    [Fact]
    public void AddFormField_DropDownWithoutOptions_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_dropdown_no_options.docx");
        var outputPath = CreateTestFilePath("test_add_dropdown_no_options_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_form_field", docPath, outputPath: outputPath,
                formFieldType: "DropDown", fieldName: "Country"));
        Assert.Contains("options array is required", ex.Message);
    }

    [Fact]
    public void AddFormField_WithInvalidFormFieldType_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_invalid_type.docx");
        var outputPath = CreateTestFilePath("test_add_invalid_type_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_form_field", docPath, outputPath: outputPath,
                formFieldType: "InvalidType", fieldName: "Field"));
        Assert.Contains("Invalid formFieldType", ex.Message);
    }

    [Fact]
    public void EditFormField_WithNonExistentField_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_nonexistent.docx");
        var outputPath = CreateTestFilePath("test_edit_nonexistent_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_form_field", docPath, outputPath: outputPath,
                fieldName: "NonExistent", value: "Value"));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void EditFormField_WithEmptyFieldName_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_edit_empty_name.docx");
        var outputPath = CreateTestFilePath("test_edit_empty_name_output.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit_form_field", docPath, outputPath: outputPath,
                fieldName: "", value: "Value"));
        Assert.Contains("fieldName is required", ex.Message);
    }

    [Fact]
    public void DeleteFormField_WithNonExistentField_ShouldDeleteZeroFields()
    {
        var docPath = CreateWordDocument("test_delete_nonexistent.docx");
        var outputPath = CreateTestFilePath("test_delete_nonexistent_output.docx");
        var result = _tool.Execute("delete_form_field", docPath, outputPath: outputPath,
            fieldName: "NonExistent");
        Assert.StartsWith("Deleted 0", result);
    }

    #endregion

    #region Session

    [Fact]
    public void InsertField_WithSessionId_ShouldInsertFieldInMemory()
    {
        var docPath = CreateWordDocument("test_session_insert.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("insert_field", sessionId: sessionId, fieldType: "DATE");
        Assert.Contains("field", result, StringComparison.OrdinalIgnoreCase);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(doc.Range.Fields.Count > 0);
    }

    [Fact]
    public void GetFields_WithSessionId_ShouldReturnFields()
    {
        var docPath = CreateWordDocument("test_session_get.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_fields", sessionId: sessionId);
        Assert.Contains("count", result);
    }

    [Fact]
    public void DeleteField_WithSessionId_ShouldDeleteInMemory()
    {
        var docPath = CreateWordDocument("test_session_delete.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        var fieldsBefore = sessionDoc.Range.Fields.Count;
        Assert.True(fieldsBefore > 0);

        _tool.Execute("delete_field", sessionId: sessionId, fieldIndex: 0);
        Assert.True(sessionDoc.Range.Fields.Count < fieldsBefore);
    }

    [Fact]
    public void UpdateAllFields_WithSessionId_ShouldUpdateInMemory()
    {
        var docPath = CreateWordDocument("test_session_update_all.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        builder.InsertField("TIME", "");
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("update_all", sessionId: sessionId);
        Assert.StartsWith("Updated", result);
    }

    [Fact]
    public void AddFormField_WithSessionId_ShouldAddFormFieldInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_form.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_form_field", sessionId: sessionId,
            formFieldType: "TextInput", fieldName: "SessionField", defaultValue: "Default");
        Assert.StartsWith("TextInput field 'SessionField' added", result);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(doc.Range.FormFields.Count > 0);
    }

    [Fact]
    public void GetFormFields_WithSessionId_ShouldReturnFormFields()
    {
        var docPath = CreateWordDocument("test_session_get_forms.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        doc.Save(docPath);

        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("get_form_fields", sessionId: sessionId);
        Assert.Contains("Name", result);
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
        var result = _tool.Execute("get_fields", docPath1, sessionId);
        Assert.Contains("TITLE", result);
        Assert.DoesNotContain("AUTHOR", result);
    }

    #endregion
}