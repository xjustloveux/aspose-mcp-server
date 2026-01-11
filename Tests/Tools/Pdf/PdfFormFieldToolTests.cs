using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Tools.Pdf;

/// <summary>
///     Integration tests for PdfFormFieldTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

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
    public void Get_WithFields_ShouldReturnFieldInfo()
    {
        var pdfPath = CreatePdfWithTextField("test_get.pdf");
        var result = _tool.Execute("get", pdfPath);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.GetProperty("count").GetInt32() > 0);
        Assert.True(json.TryGetProperty("items", out _));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pdfPath = CreateTestPdf($"test_case_{operation}.pdf");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.pdf");
        var result = _tool.Execute(operation, pdfPath, outputPath: outputPath,
            fieldType: "TextField", fieldName: $"field_{operation}",
            pageIndex: 1, x: 100, y: 700, width: 200, height: 20);
        Assert.StartsWith("Added", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pdfPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

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
        var result = _tool.Execute("add", sessionId: sessionId,
            fieldType: "TextField", fieldName: "sessionField",
            pageIndex: 1, x: 100, y: 700, width: 200, height: 20);
        Assert.StartsWith("Added", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteFromSession()
    {
        var pdfPath = CreatePdfWithTextField("test_session_delete.pdf", "fieldToDelete");
        var sessionId = OpenSession(pdfPath);
        var result = _tool.Execute("delete", sessionId: sessionId, fieldName: "fieldToDelete");
        Assert.StartsWith("Deleted", result);
        Assert.Contains("session", result);
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
