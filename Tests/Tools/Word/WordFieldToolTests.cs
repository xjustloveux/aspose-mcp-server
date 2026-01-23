using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Word.Field;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordFieldTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class WordFieldToolTests : WordTestBase
{
    private readonly WordFieldTool _tool;

    public WordFieldToolTests()
    {
        _tool = new WordFieldTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void InsertField_ShouldInsertFieldAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_insert_field.docx");
        var outputPath = CreateTestFilePath("test_insert_field_output.docx");
        _tool.Execute("insert_field", docPath, outputPath: outputPath, fieldType: "DATE");
        var doc = new Document(outputPath);
        var fields = doc.Range.Fields;
        Assert.True(fields.Count > 0);
    }

    [Fact]
    public void GetFields_ShouldReturnFieldsFromFile()
    {
        var docPath = CreateWordDocument("test_get_fields.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        builder.InsertField("TIME", "");
        doc.Save(docPath);
        var result = _tool.Execute("get_fields", docPath);
        var data = GetResultData<GetFieldsWordResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.Fields, f => f.Type == "FieldDate");
    }

    [Fact]
    public void UpdateField_ShouldUpdateFieldAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_update_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_field_output.docx");
        var result = _tool.Execute("update_field", docPath, outputPath: outputPath, fieldIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("updated", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void UpdateAllFields_ShouldUpdateAllFieldsAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_update_all.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        builder.InsertField("TIME", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_update_all_output.docx");
        var result = _tool.Execute("update_all", docPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Updated", data.Message);
    }

    [Fact]
    public void EditField_ShouldEditFieldAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_edit_field.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_edit_field_output.docx");
        var result = _tool.Execute("edit_field", docPath, outputPath: outputPath,
            fieldIndex: 0, fieldCode: "DATE \\@ \"yyyy-MM-dd\"");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("edited", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteField_ShouldDeleteFieldAndPersistToFile()
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
    public void GetFieldDetail_ShouldReturnFieldDetailFromFile()
    {
        var docPath = CreateWordDocument("test_get_detail.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertField("DATE", "");
        doc.Save(docPath);
        var result = _tool.Execute("get_field_detail", docPath, fieldIndex: 0);
        var data = GetResultData<GetFieldDetailWordResult>(result);
        Assert.NotNull(data.Type);
        Assert.NotNull(data.Code);
    }

    [Fact]
    public void AddFormField_TextInput_ShouldAddFieldAndPersistToFile()
    {
        var docPath = CreateWordDocument("test_add_text.docx");
        var outputPath = CreateTestFilePath("test_add_text_output.docx");
        _tool.Execute("add_form_field", docPath, outputPath: outputPath,
            formFieldType: "TextInput", fieldName: "Name", defaultValue: "Default");
        var doc = new Document(outputPath);
        Assert.NotNull(doc.Range.FormFields["Name"]);
    }

    [Fact]
    public void GetFormFields_ShouldReturnFormFieldsFromFile()
    {
        var docPath = CreateWordDocument("test_get_forms.docx");
        var doc = new Document(docPath);
        var builder = new DocumentBuilder(doc);
        builder.InsertTextInput("Name", TextFormFieldType.Regular, "", "", 0);
        builder.InsertCheckBox("Accept", false, 0);
        doc.Save(docPath);
        var result = _tool.Execute("get_form_fields", docPath);
        var data = GetResultData<GetFormFieldsWordResult>(result);
        Assert.True(data.Count > 0);
        Assert.Contains(data.FormFields, f => f.Name == "Name");
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("INSERT_FIELD")]
    [InlineData("Insert_Field")]
    [InlineData("insert_field")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_case_{operation.Replace("_", "")}.docx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, fieldType: "DATE");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Field inserted successfully", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_unknown_op.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_fields"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void InsertField_WithSessionId_ShouldInsertFieldInMemory()
    {
        var docPath = CreateWordDocument("test_session_insert.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("insert_field", sessionId: sessionId, fieldType: "DATE");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("field", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<GetFieldsWordResult>(result);
        Assert.True(data.Count > 0);
        var output = GetResultOutput<GetFieldsWordResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Updated", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void AddFormField_WithSessionId_ShouldAddFormFieldInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_form.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_form_field", sessionId: sessionId,
            formFieldType: "TextInput", fieldName: "SessionField", defaultValue: "Default");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("TextInput field 'SessionField' added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<GetFormFieldsWordResult>(result);
        Assert.Contains(data.FormFields, f => f.Name == "Name");
        var output = GetResultOutput<GetFormFieldsWordResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<GetFieldsWordResult>(result);
        Assert.Contains(data.Fields, f => f.Type == "FieldTitle");
        Assert.DoesNotContain(data.Fields, f => f.Type == "FieldAuthor");
    }

    #endregion
}
