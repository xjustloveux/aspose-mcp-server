using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfFormFieldToolTests : PdfTestBase
{
    private readonly PdfFormFieldTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task AddTextField_ShouldAddTextField()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_text_field.pdf");
        var outputPath = CreateTestFilePath("test_add_text_field_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["fieldType"] = "TextField",
            ["fieldName"] = "testField",
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 700,
            ["width"] = 200,
            ["height"] = 20
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var document = new Document(outputPath);
        var form = document.Form;
        Assert.True(form.Count > 0, "Form should contain at least one field");
    }

    [Fact]
    public async Task GetFormFields_ShouldReturnAllFields()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_form_fields.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var textField = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "testField"
        };
        document.Form.Add(textField);
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Field", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteFormField_ShouldDeleteField()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["fieldName"] = "fieldToDelete"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDocument = new Document(outputPath);
        var fieldsAfter = resultDocument.Form.Count;
        Assert.True(fieldsAfter < fieldsBefore,
            $"Field should be deleted. Before: {fieldsBefore}, After: {fieldsAfter}");
    }

    [Fact]
    public async Task EditFormField_ShouldEditField()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["fieldName"] = "fieldToEdit",
            ["value"] = "Updated Value"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDocument = new Document(outputPath);
        var field = resultDocument.Form["fieldToEdit"] as TextBoxField;
        Assert.NotNull(field);
        Assert.Equal("Updated Value", field.Value);
    }

    [Fact]
    public async Task AddCheckBox_ShouldAddCheckBoxField()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_checkbox.pdf");
        var outputPath = CreateTestFilePath("test_add_checkbox_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["fieldType"] = "CheckBox",
            ["fieldName"] = "testCheckBox",
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 700,
            ["width"] = 20,
            ["height"] = 20
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var document = new Document(outputPath);
        Assert.True(document.Form.Count > 0, "Form should contain checkbox field");
    }

    [Fact]
    public async Task EditCheckBox_ShouldUpdateCheckedState()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["fieldName"] = "checkBoxToEdit",
            ["checkedValue"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultDocument = new Document(outputPath);
        var field = resultDocument.Form["checkBoxToEdit"] as CheckboxField;
        Assert.NotNull(field);
        Assert.True(field.Checked);
    }

    [Fact]
    public async Task AddFormField_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["fieldType"] = "TextBox",
            ["fieldName"] = "testField",
            ["pageIndex"] = 99,
            ["x"] = 100,
            ["y"] = 700,
            ["width"] = 200,
            ["height"] = 20
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task AddFormField_WithDuplicateName_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_duplicate_name.pdf");
        var document = new Document(pdfPath);
        var page = document.Pages[1];
        var existingField = new TextBoxField(page, new Rectangle(100, 700, 300, 720))
        {
            PartialName = "existingField"
        };
        document.Form.Add(existingField);
        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["fieldType"] = "TextBox",
            ["fieldName"] = "existingField",
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 600,
            ["width"] = 200,
            ["height"] = 20
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("already exists", exception.Message);
    }

    [Fact]
    public async Task DeleteFormField_WithNonExistentField_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_nonexistent.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["fieldName"] = "nonExistentField"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public async Task EditFormField_WithNonExistentField_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_edit_nonexistent.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pdfPath,
            ["fieldName"] = "nonExistentField",
            ["value"] = "test"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public async Task GetFormFields_WithLimit_ShouldRespectLimit()
    {
        // Arrange
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

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath,
            ["limit"] = 3
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 3", result);
        Assert.Contains("\"totalCount\": 5", result);
        Assert.Contains("\"truncated\": true", result);
    }

    [Fact]
    public async Task Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pdfPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task GetFormFields_WithNoFields_ShouldReturnEmptyResult()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_empty.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No form fields found", result);
    }
}