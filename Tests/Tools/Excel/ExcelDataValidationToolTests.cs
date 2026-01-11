using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelDataValidationTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelDataValidationToolTests : ExcelTestBase
{
    private readonly ExcelDataValidationTool _tool;

    public ExcelDataValidationToolTests()
    {
        _tool = new ExcelDataValidationTool(SessionManager);
    }

    private string CreateWorkbookWithValidation(string fileName)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var idx = worksheet.Validations.Add(area);
        worksheet.Validations[idx].Type = ValidationType.List;
        worksheet.Validations[idx].Formula1 = "1,2,3";
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_WithList_ShouldAddListValidation()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_list.xlsx");
        var outputPath = CreateTestFilePath("test_add_list_output.xlsx");
        var result = _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "List",
            formula1: "Option1,Option2,Option3", outputPath: outputPath);
        Assert.StartsWith("Data validation added", result);
        using var workbook = new Workbook(outputPath);
        var validation = workbook.Worksheets[0].Validations[^1];
        Assert.Equal(ValidationType.List, validation.Type);
    }

    [Fact]
    public void Get_ShouldReturnValidationInfo()
    {
        var workbookPath = CreateWorkbookWithValidation("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        Assert.Contains("count", result);
    }

    [Fact]
    public void Edit_ShouldEditValidation()
    {
        var workbookPath = CreateWorkbookWithValidation("test_edit.xlsx");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, validationIndex: 0, validationType: "WholeNumber",
            formula1: "0", formula2: "100", outputPath: outputPath);
        Assert.StartsWith("Edited data validation #0", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(ValidationType.WholeNumber, workbook.Worksheets[0].Validations[0].Type);
    }

    [Fact]
    public void Delete_ShouldDeleteValidation()
    {
        var workbookPath = CreateWorkbookWithValidation("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, validationIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Deleted data validation #0", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].Validations);
    }

    [Fact]
    public void SetMessages_ShouldSetInputAndErrorMessage()
    {
        var workbookPath = CreateWorkbookWithValidation("test_set_messages.xlsx");
        var outputPath = CreateTestFilePath("test_set_messages_output.xlsx");
        var result = _tool.Execute("set_messages", workbookPath, validationIndex: 0,
            inputMessage: "Please select a value", errorMessage: "Invalid value selected", outputPath: outputPath);
        Assert.StartsWith("Updated data validation #0", result);
        using var workbook = new Workbook(outputPath);
        var validation = workbook.Worksheets[0].Validations[0];
        Assert.Equal("Please select a value", validation.InputMessage);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:A10", validationType: "List", formula1: "A,B,C",
            outputPath: outputPath);
        Assert.StartsWith("Data validation added", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, range: "A1:A10", validationType: "List",
            formula1: "SessionOpt1,SessionOpt2");
        Assert.StartsWith("Data validation added", result);
        Assert.Contains("session", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Validations.Count > 0);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithValidation("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("count", result);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithValidation("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("delete", sessionId: sessionId, validationIndex: 0);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Validations);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateExcelWorkbook("test_path_file.xlsx");
        var sessionWorkbook = CreateWorkbookWithValidation("test_session_file.xlsx");
        using (var wb = new Workbook(sessionWorkbook))
        {
            wb.Worksheets[0].Name = "SessionSheet";
            wb.Save(sessionWorkbook);
        }

        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        Assert.Contains("SessionSheet", result);
    }

    #endregion
}
