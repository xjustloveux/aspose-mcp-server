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
}