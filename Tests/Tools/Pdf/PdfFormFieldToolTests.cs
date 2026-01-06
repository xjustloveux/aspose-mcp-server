using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

public class PdfFormFieldToolTests : PdfTestBase
{
    private readonly PdfFormFieldTool _tool;

    public PdfFormFieldToolTests()
    {
        _tool = new PdfFormFieldTool(SessionManager);
    }

    private string CreateTestPdf(string fileName, int pageCount = 1)
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreatePdfWithTextField(string fileName, string fieldName = "testField")
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        var textField = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = fieldName,
            Value = "Test Value"
        };
        document.Form.Add(textField);
        document.Save(filePath);
        return filePath;
    }

    private string CreatePdfWithCheckBox(string fileName, string fieldName = "testCheckBox")
    {
        var filePath = CreateTestFilePath(fileName);
        using var document = new Document();
        var page = document.Pages.Add();
        var checkBox = new CheckboxField(page, new Rectangle(100, 700, 120, 720))
        {
            PartialName = fieldName,
            Checked = false
        };
        document.Form.Add(checkBox);
        document.Save(filePath);
        return filePath;
    }

    #region General

    [Fact]
    public void Add_TextField_ShouldAddField()
    {
        var pdfPath = CreateTestPdf("test_add_text.pdf");
        var outputPath = CreateTestFilePath("test_add_text_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            fieldType: "TextField", fieldName: "newField",
            pageIndex: 1, x: 100, y: 700, width: 200, height: 20);
        Assert.StartsWith("Added", result);
        using var document = new Document(outputPath);
        Assert.True(document.Form.Count > 0);
    }

    [Fact]
    public void Add_TextBox_ShouldAddField()
    {
        var pdfPath = CreateTestPdf("test_add_textbox.pdf");
        var outputPath = CreateTestFilePath("test_add_textbox_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            fieldType: "TextBox", fieldName: "textBoxField",
            pageIndex: 1, x: 100, y: 700, width: 200, height: 20);
        Assert.StartsWith("Added", result);
        using var document = new Document(outputPath);
        Assert.True(document.Form.Count > 0);
    }

    [Fact]
    public void Add_CheckBox_ShouldAddField()
    {
        var pdfPath = CreateTestPdf("test_add_checkbox.pdf");
        var outputPath = CreateTestFilePath("test_add_checkbox_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            fieldType: "CheckBox", fieldName: "checkBoxField",
            pageIndex: 1, x: 100, y: 700, width: 20, height: 20);
        Assert.StartsWith("Added", result);
        using var document = new Document(outputPath);
        Assert.True(document.Form.Count > 0);
    }

    [Fact]
    public void Add_RadioButton_ShouldAddField()
    {
        var pdfPath = CreateTestPdf("test_add_radio.pdf");
        var outputPath = CreateTestFilePath("test_add_radio_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            fieldType: "RadioButton", fieldName: "radioField",
            pageIndex: 1, x: 100, y: 700, width: 20, height: 20);
        Assert.StartsWith("Added", result);
        using var document = new Document(outputPath);
        Assert.True(document.Form.Count > 0);
    }

    [Fact]
    public void Add_RadioButton_WithDefaultValue_ShouldUseAsOptionName()
    {
        var pdfPath = CreateTestPdf("test_add_radio_option.pdf");
        var outputPath = CreateTestFilePath("test_add_radio_option_output.pdf");
        var result = _tool.Execute("add", pdfPath, outputPath: outputPath,
            fieldType: "RadioButton", fieldName: "radioField",
            pageIndex: 1, x: 100, y: 700, width: 20, height: 20,
            defaultValue: "CustomOption");
        Assert.StartsWith("Added", result);
        using var document = new Document(outputPath);
        var field = document.Form["radioField"] as RadioButtonField;
        Assert.NotNull(field);
    }

    [Fact]
    public void Add_WithDefaultValue_ShouldSetValue()
    {
        var pdfPath = CreateTestPdf("test_add_default.pdf");
        var outputPath = CreateTestFilePath("test_add_default_output.pdf");
        _tool.Execute("add", pdfPath, outputPath: outputPath,
            fieldType: "TextField", fieldName: "fieldWithDefault",
            pageIndex: 1, x: 100, y: 700, width: 200, height: 20,
            defaultValue: "Default Text");
        using var document = new Document(outputPath);
        var field = document.Form["fieldWithDefault"] as TextBoxField;
        Assert.NotNull(field);
        Assert.Equal("Default Text", field.Value);
    }

    [Fact]
    public void Delete_ShouldDeleteField()
    {
        var pdfPath = CreatePdfWithTextField("test_delete.pdf", "fieldToDelete");
        var outputPath = CreateTestFilePath("test_delete_output.pdf");
        var result = _tool.Execute("delete", pdfPath, outputPath: outputPath,
            fieldName: "fieldToDelete");
        Assert.StartsWith("Deleted", result);
        using var document = new Document(outputPath);
        Assert.Empty(document.Form);
    }

    [Fact]
    public void Edit_TextField_ShouldUpdateValue()
    {
        var pdfPath = CreatePdfWithTextField("test_edit_text.pdf", "fieldToEdit");
        var outputPath = CreateTestFilePath("test_edit_text_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            fieldName: "fieldToEdit", value: "Updated Value");
        Assert.StartsWith("Edited", result);
        using var document = new Document(outputPath);
        var field = document.Form["fieldToEdit"] as TextBoxField;
        Assert.NotNull(field);
        Assert.Equal("Updated Value", field.Value);
    }

    [Fact]
    public void Edit_CheckBox_ShouldUpdateCheckedState()
    {
        var pdfPath = CreatePdfWithCheckBox("test_edit_checkbox.pdf", "checkToEdit");
        var outputPath = CreateTestFilePath("test_edit_checkbox_output.pdf");
        var result = _tool.Execute("edit", pdfPath, outputPath: outputPath,
            fieldName: "checkToEdit", checkedValue: true);
        Assert.StartsWith("Edited", result);
        using var document = new Document(outputPath);
        var field = document.Form["checkToEdit"] as CheckboxField;
        Assert.NotNull(field);
        Assert.True(field.Checked);
    }

    [Fact]
    public void Get_WithFields_ShouldReturnFieldInfo()
    {
        var pdfPath = CreatePdfWithTextField("test_get.pdf");
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
        Assert.True(json.TryGetProperty("items", out _));
    }

    [Fact]
    public void Get_WithNoFields_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(0, json.GetProperty("count").GetInt32());
        Assert.Contains("No form fields found", result);
    }

    [Fact]
    public void Get_WithLimit_ShouldRespectLimit()
    {
        var pdfPath = CreateTestFilePath("test_get_limit.pdf");
        using (var document = new Document())
        {
            var page = document.Pages.Add();
            for (var i = 0; i < 5; i++)
            {
                var field = new TextBoxField(page, new Rectangle(100, 700 - i * 30, 300, 720 - i * 30))
                {
                    PartialName = $"field{i}"
                };
                document.Form.Add(field);
            }

            document.Save(pdfPath);
        }

        var result = _tool.Execute("get", pdfPath, limit: 3);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.Equal(3, json.GetProperty("count").GetInt32());
        Assert.Equal(5, json.GetProperty("totalCount").GetInt32());
        Assert.True(json.GetProperty("truncated").GetBoolean());
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            fieldType: "TextField", fieldName: $"field_{operation}",
            pageIndex: 1, x: 100, y: 700, width: 200, height: 20);
        Assert.StartsWith("Added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_get_{operation}.pdf");
        var result = _tool.Execute(operation, pdfPath);
        Assert.Contains("count", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_invalid_page.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, fieldType: "TextBox", fieldName: "field",
                pageIndex: 99, x: 100, y: 700, width: 200, height: 20));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Add_WithMissingFieldName_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_name.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, fieldType: "TextField",
                pageIndex: 1, x: 100, y: 700, width: 200, height: 20));
        Assert.Contains("fieldName is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingFieldType_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_type.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, fieldName: "field",
                pageIndex: 1, x: 100, y: 700, width: 200, height: 20));
        Assert.Contains("fieldType is required", ex.Message);
    }

    [Fact]
    public void Add_WithUnknownFieldType_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_unknown_type.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, fieldType: "UnknownType", fieldName: "field",
                pageIndex: 1, x: 100, y: 700, width: 200, height: 20));
        Assert.Contains("Unknown field type", ex.Message);
    }

    [Fact]
    public void Add_WithDuplicateName_ShouldThrowArgumentException()
    {
        var pdfPath = CreatePdfWithTextField("test_add_duplicate.pdf", "existingField");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, fieldType: "TextBox", fieldName: "existingField",
                pageIndex: 1, x: 100, y: 600, width: 200, height: 20));
        Assert.Contains("already exists", ex.Message);
    }

    [Fact]
    public void Add_WithMissingCoordinates_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_add_no_coords.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pdfPath, fieldType: "TextField", fieldName: "field",
                pageIndex: 1, width: 200, height: 20));
        Assert.Contains("x is required", ex.Message);
    }

    [Fact]
    public void Delete_WithNonExistentField_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_nonexistent.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath, fieldName: "nonExistent"));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Delete_WithMissingFieldName_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_no_name.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pdfPath));
        Assert.Contains("fieldName is required", ex.Message);
    }

    [Fact]
    public void Edit_WithNonExistentField_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_edit_nonexistent.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, fieldName: "nonExistent", value: "test"));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Edit_WithMissingFieldName_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_edit_no_name.pdf");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pdfPath, value: "test"));
        Assert.Contains("fieldName is required", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreatePdfWithTextField("test_session_get.pdf", "sessionField");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("sessionField", result);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var countBefore = docBefore.Form.Count;
        var result = _tool.Execute("add", sessionId: sessionId,
            fieldType: "TextField", fieldName: "sessionField",
            pageIndex: 1, x: 100, y: 700, width: 200, height: 20);
        Assert.StartsWith("Added", result);
        Assert.Contains("session", result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docAfter.Form.Count > countBefore);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithTextField("test_session_delete.pdf", "fieldToDelete");
        var sessionId = OpenSession(pdfPath);
        var docBefore = SessionManager.GetDocument<Document>(sessionId);
        var countBefore = docBefore.Form.Count;
        Assert.True(countBefore > 0);
        var result = _tool.Execute("delete", sessionId: sessionId, fieldName: "fieldToDelete");
        Assert.StartsWith("Deleted", result);
        Assert.Contains("session", result);
        var docAfter = SessionManager.GetDocument<Document>(sessionId);
        Assert.True(docAfter.Form.Count < countBefore);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInSession()
    {
        var pdfPath = CreatePdfWithTextField("test_session_edit.pdf", "fieldToEdit");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("edit", sessionId: sessionId,
            fieldName: "fieldToEdit", value: "Session Updated");
        Assert.StartsWith("Edited", result);
        Assert.Contains("session", result);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        var field = doc.Form["fieldToEdit"] as TextBoxField;
        Assert.NotNull(field);
        Assert.Equal("Session Updated", field.Value);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pdfPath1 = CreateTestPdf("test_path_form.pdf");
        var pdfPath2 = CreatePdfWithTextField("test_session_form.pdf", "sessionOnlyField");
        var sessionId = OpenSession(pdfPath2);
        var result = _tool.Execute("get", pdfPath1, sessionId);
        Assert.Contains("sessionOnlyField", result);
    }

    #endregion
}