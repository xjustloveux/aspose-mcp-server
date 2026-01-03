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

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddTextField_ShouldAddTextField()
    {
        var pdfPath = CreateTestPdf("test_add_text_field.pdf");
        var outputPath = CreateTestFilePath("test_add_text_field_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            fieldType: "TextField",
            fieldName: "testField",
            pageIndex: 1,
            x: 100,
            y: 700,
            width: 200,
            height: 20);
        var document = new Document(outputPath);
        var form = document.Form;
        Assert.True(form.Count > 0, "Form should contain at least one field");
    }

    [Fact]
    public void GetFormFields_ShouldReturnAllFields()
    {
        var pdfPath = CreateTestPdf("test_get_form_fields.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var textField = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "testField"
        };
        document.Form.Add(textField);
        document.Save(pdfPath);
        var result = _tool.Execute("get", pdfPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Field", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteFormField_ShouldDeleteField()
    {
        var pdfPath = CreateTestPdf("test_delete_form_field.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var textField = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "fieldToDelete"
        };
        document.Form.Add(textField);
        document.Save(pdfPath);

        var fieldsBefore = document.Form.Count;
        Assert.True(fieldsBefore > 0, "Field should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_form_field_output.pdf");
        _tool.Execute(
            "delete",
            pdfPath,
            outputPath: outputPath,
            fieldName: "fieldToDelete");
        var resultDocument = new Document(outputPath);
        var fieldsAfter = resultDocument.Form.Count;
        Assert.True(fieldsAfter < fieldsBefore,
            $"Field should be deleted. Before: {fieldsBefore}, After: {fieldsAfter}");
    }

    [Fact]
    public void EditFormField_ShouldEditField()
    {
        var pdfPath = CreateTestPdf("test_edit_form_field.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var textField = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "fieldToEdit",
            Value = "Original Value"
        };
        document.Form.Add(textField);
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_edit_form_field_output.pdf");
        _tool.Execute(
            "edit",
            pdfPath,
            outputPath: outputPath,
            fieldName: "fieldToEdit",
            value: "Updated Value");
        var resultDocument = new Document(outputPath);
        var field = resultDocument.Form["fieldToEdit"] as TextBoxField;
        Assert.NotNull(field);
        Assert.Equal("Updated Value", field.Value);
    }

    [Fact]
    public void AddCheckBox_ShouldAddCheckBoxField()
    {
        var pdfPath = CreateTestPdf("test_add_checkbox.pdf");
        var outputPath = CreateTestFilePath("test_add_checkbox_output.pdf");
        _tool.Execute(
            "add",
            pdfPath,
            outputPath: outputPath,
            fieldType: "CheckBox",
            fieldName: "testCheckBox",
            pageIndex: 1,
            x: 100,
            y: 700,
            width: 20,
            height: 20);
        var document = new Document(outputPath);
        Assert.True(document.Form.Count > 0, "Form should contain checkbox field");
    }

    [Fact]
    public void EditCheckBox_ShouldUpdateCheckedState()
    {
        var pdfPath = CreateTestPdf("test_edit_checkbox.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var checkBox = new CheckboxField(page, new Rectangle(100, 700, 120, 720))
        {
            PartialName = "checkBoxToEdit",
            Checked = false
        };
        document.Form.Add(checkBox);
        document.Save(pdfPath);

        var outputPath = CreateTestFilePath("test_edit_checkbox_output.pdf");
        _tool.Execute(
            "edit",
            pdfPath,
            outputPath: outputPath,
            fieldName: "checkBoxToEdit",
            checkedValue: true);
        var resultDocument = new Document(outputPath);
        var field = resultDocument.Form["checkBoxToEdit"] as CheckboxField;
        Assert.NotNull(field);
        Assert.True(field.Checked);
    }

    [Fact]
    public void AddFormField_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_invalid_page.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            fieldType: "TextBox",
            fieldName: "testField",
            pageIndex: 99,
            x: 100,
            y: 700,
            width: 200,
            height: 20));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public void AddFormField_WithDuplicateName_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_duplicate_name.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var existingField = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "existingField"
        };
        document.Form.Add(existingField);
        document.Save(pdfPath);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            fieldType: "TextBox",
            fieldName: "existingField",
            pageIndex: 1,
            x: 100,
            y: 600,
            width: 200,
            height: 20));
        Assert.Contains("already exists", exception.Message);
    }

    [Fact]
    public void DeleteFormField_WithNonExistentField_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_delete_nonexistent.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            pdfPath,
            fieldName: "nonExistentField"));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void EditFormField_WithNonExistentField_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_edit_nonexistent.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            pdfPath,
            fieldName: "nonExistentField",
            value: "test"));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void GetFormFields_WithLimit_ShouldRespectLimit()
    {
        var pdfPath = CreateTestPdf("test_get_with_limit.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        for (var i = 0; i < 5; i++)
        {
            var field = new TextBoxField(page, new Rectangle(100, 700 - i * 30, 300, 720 - i * 30))
            {
                PartialName = $"field{i}"
            };
            document.Form.Add(field);
        }

        document.Save(pdfPath);
        var result = _tool.Execute("get", pdfPath, limit: 3);
        Assert.Contains("\"count\": 3", result);
        Assert.Contains("\"totalCount\": 5", result);
        Assert.Contains("\"truncated\": true", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void GetFormFields_WithNoFields_ShouldReturnEmptyResult()
    {
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var result = _tool.Execute("get", pdfPath);
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No form fields found", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithInvalidOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_exception_unknown.pdf");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute("invalid_operation", pdfPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void AddFormField_WithMissingFieldName_ShouldThrow()
    {
        var pdfPath = CreateTestPdf("test_missing_field_name.pdf");

        // Act & Assert - missing fieldName
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            pdfPath,
            fieldType: "TextField",
            pageIndex: 1,
            x: 100,
            y: 700,
            width: 200,
            height: 20));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetFormFields_WithSessionId_ShouldGetFromSession()
    {
        var pdfPath = CreateTestPdf("test_session_get.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var textField = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "sessionField"
        };
        document.Form.Add(textField);
        document.Save(pdfPath);

        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("sessionField", result);
    }

    [Fact]
    public void AddFormField_WithSessionId_ShouldAddToSession()
    {
        var pdfPath = CreateTestPdf("test_session_add.pdf");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            fieldType: "TextField",
            fieldName: "newSessionField",
            pageIndex: 1,
            x: 100,
            y: 700,
            width: 200,
            height: 20);
        Assert.Contains("Added", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void AddFormField_WithSessionId_ShouldModifyInMemory()
    {
        var pdfPath = CreateTestPdf("test_session_memory.pdf");
        var sessionId = OpenSession(pdfPath);
        _tool.Execute(
            "add",
            sessionId: sessionId,
            fieldType: "TextField",
            fieldName: "inMemoryField",
            pageIndex: 1,
            x: 100,
            y: 700,
            width: 200,
            height: 20);

        // Assert - verify in-memory changes
        var document = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(document);
        Assert.True(document.Form.Count > 0);
    }

    #endregion
}