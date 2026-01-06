using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelDataValidationToolTests : ExcelTestBase
{
    private readonly ExcelDataValidationTool _tool;

    public ExcelDataValidationToolTests()
    {
        _tool = new ExcelDataValidationTool(SessionManager);
    }

    private string CreateWorkbookWithValidation(string fileName, ValidationType type = ValidationType.List,
        string formula = "1,2,3")
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var idx = worksheet.Validations.Add(area);
        worksheet.Validations[idx].Type = type;
        worksheet.Validations[idx].Formula1 = formula;
        workbook.Save(path);
        return path;
    }

    #region General

    [Fact]
    public void Add_WithList_ShouldAddListValidation()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_list.xlsx");
        var outputPath = CreateTestFilePath("test_add_list_output.xlsx");
        var result = _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "List",
            formula1: "Option1,Option2,Option3", outputPath: outputPath);
        Assert.StartsWith("Data validation added", result);
        Assert.Contains("index: 0", result);
        using var workbook = new Workbook(outputPath);
        var validation = workbook.Worksheets[0].Validations[^1];
        Assert.Equal(ValidationType.List, validation.Type);
        Assert.True(validation.InCellDropDown);
    }

    [Fact]
    public void Add_WithList_InCellDropDownFalse_ShouldDisableDropdown()
    {
        var workbookPath = CreateExcelWorkbook("test_add_list_no_dropdown.xlsx");
        var outputPath = CreateTestFilePath("test_add_list_no_dropdown_output.xlsx");
        _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "List", formula1: "1,2,3",
            inCellDropDown: false, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var validation = workbook.Worksheets[0].Validations[^1];
        Assert.Equal(ValidationType.List, validation.Type);
        Assert.False(validation.InCellDropDown);
    }

    [Fact]
    public void Add_WithWholeNumber_ShouldAddNumberValidation()
    {
        var workbookPath = CreateExcelWorkbook("test_add_number.xlsx");
        var outputPath = CreateTestFilePath("test_add_number_output.xlsx");
        _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "WholeNumber", formula1: "0",
            formula2: "100", outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var validation = workbook.Worksheets[0].Validations[^1];
        Assert.Equal(ValidationType.WholeNumber, validation.Type);
        Assert.Equal(OperatorType.Between, validation.Operator);
    }

    [Fact]
    public void Add_WithMessages_ShouldSetMessages()
    {
        var workbookPath = CreateExcelWorkbook("test_add_messages.xlsx");
        var outputPath = CreateTestFilePath("test_add_messages_output.xlsx");
        _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "List", formula1: "Yes,No",
            inputMessage: "Please select Yes or No", errorMessage: "Invalid selection", outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var validation = workbook.Worksheets[0].Validations[^1];
        Assert.Equal("Please select Yes or No", validation.InputMessage);
        Assert.Equal("Invalid selection", validation.ErrorMessage);
        Assert.True(validation.ShowInput);
        Assert.True(validation.ShowError);
    }

    [Theory]
    [InlineData("WholeNumber")]
    [InlineData("Decimal")]
    [InlineData("Date")]
    [InlineData("Time")]
    [InlineData("TextLength")]
    [InlineData("Custom")]
    public void Add_AllValidationTypes_ShouldWork(string validationType)
    {
        var workbookPath = CreateExcelWorkbook($"test_type_{validationType}.xlsx");
        var outputPath = CreateTestFilePath($"test_type_{validationType}_output.xlsx");
        var result = _tool.Execute("add", workbookPath, range: "A1:A10", validationType: validationType, formula1: "1",
            outputPath: outputPath);
        Assert.StartsWith("Data validation added", result);
        Assert.Contains($"type: {validationType}", result);
    }

    [Theory]
    [InlineData("Between", "100")]
    [InlineData("Equal", null)]
    [InlineData("NotEqual", null)]
    [InlineData("GreaterThan", null)]
    [InlineData("LessThan", null)]
    [InlineData("GreaterOrEqual", null)]
    [InlineData("LessOrEqual", null)]
    public void Add_AllOperatorTypes_ShouldWork(string operatorType, string? formula2)
    {
        var workbookPath = CreateExcelWorkbook($"test_op_{operatorType}.xlsx");
        var outputPath = CreateTestFilePath($"test_op_{operatorType}_output.xlsx");
        var result = _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "WholeNumber",
            operatorType: operatorType, formula1: "0", formula2: formula2, outputPath: outputPath);
        Assert.StartsWith("Data validation added", result);
    }

    [Fact]
    public void Get_ShouldReturnValidationInfo()
    {
        var workbookPath = CreateWorkbookWithValidation("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.True(json.RootElement.GetProperty("items").GetArrayLength() > 0);
    }

    [Fact]
    public void Get_EmptyWorksheet_ShouldReturnEmptyList()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(0, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void Edit_ShouldEditValidation()
    {
        var workbookPath = CreateWorkbookWithValidation("test_edit.xlsx");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, validationIndex: 0, validationType: "WholeNumber",
            formula1: "0", formula2: "100", outputPath: outputPath);
        Assert.StartsWith("Edited data validation #0", result);
        Assert.Contains("Type=WholeNumber", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(ValidationType.WholeNumber, workbook.Worksheets[0].Validations[0].Type);
    }

    [Fact]
    public void Edit_ChangeOperatorType_ShouldUpdateOperator()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_operator.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
            var idx = wb.Worksheets[0].Validations.Add(area);
            wb.Worksheets[0].Validations[idx].Type = ValidationType.WholeNumber;
            wb.Worksheets[0].Validations[idx].Formula1 = "0";
            wb.Worksheets[0].Validations[idx].Operator = OperatorType.Equal;
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_operator_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, validationIndex: 0, operatorType: "GreaterOrEqual",
            outputPath: outputPath);
        Assert.StartsWith("Edited data validation #0", result);
        Assert.Contains("Operator=GreaterOrEqual", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(OperatorType.GreaterOrEqual, workbook.Worksheets[0].Validations[0].Operator);
    }

    [Fact]
    public void Edit_ChangeInCellDropDown_ShouldUpdateDropdown()
    {
        var workbookPath = CreateWorkbookWithValidation("test_edit_dropdown.xlsx");
        var outputPath = CreateTestFilePath("test_edit_dropdown_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, validationIndex: 0, inCellDropDown: false,
            outputPath: outputPath);
        Assert.StartsWith("Edited data validation #0", result);
        Assert.Contains("InCellDropDown=False", result);
        using var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].Validations[0].InCellDropDown);
    }

    [Fact]
    public void Delete_ShouldDeleteValidation()
    {
        var workbookPath = CreateWorkbookWithValidation("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, validationIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Deleted data validation #0", result);
        Assert.Contains("remaining: 0", result);
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
        Assert.Contains("InputMessage=Please select a value", result);
        Assert.Contains("ErrorMessage=Invalid value selected", result);
        using var workbook = new Workbook(outputPath);
        var validation = workbook.Worksheets[0].Validations[0];
        Assert.Equal("Please select a value", validation.InputMessage);
        Assert.Equal("Invalid value selected", validation.ErrorMessage);
    }

    [Fact]
    public void SetMessages_ClearMessage_ShouldClearAndDisableShow()
    {
        var workbookPath = CreateExcelWorkbook("test_clear_messages.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
            var idx = wb.Worksheets[0].Validations.Add(area);
            wb.Worksheets[0].Validations[idx].InputMessage = "Old message";
            wb.Worksheets[0].Validations[idx].ShowInput = true;
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_clear_messages_output.xlsx");
        _tool.Execute("set_messages", workbookPath, validationIndex: 0, inputMessage: "", outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        var validation = workbook.Worksheets[0].Validations[0];
        Assert.True(string.IsNullOrEmpty(validation.InputMessage));
        Assert.False(validation.ShowInput);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:A10", validationType: "List", formula1: "A,B,C",
            outputPath: outputPath);
        Assert.StartsWith("Data validation added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("count", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var workbookPath = CreateWorkbookWithValidation($"test_case_delete_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, validationIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Deleted data validation #0", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, validationType: "List", formula1: "A,B,C"));
        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingValidationType_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_type.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:A10", formula1: "A,B,C"));
        Assert.Contains("validationType is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingFormula1_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_formula.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "List"));
        Assert.Contains("formula1 is required", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidValidationType_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_type.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "InvalidType", formula1: "test"));
        Assert.Contains("Unsupported validation type", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidOperatorType_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_operator.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:A10", validationType: "WholeNumber",
                operatorType: "InvalidOperator", formula1: "0"));
        Assert.Contains("Unsupported operator type", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, validationIndex: 99, formula1: "test"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath, validationIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void SetMessages_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_set_messages_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_messages", workbookPath, validationIndex: 99, inputMessage: "test"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

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
        Assert.Equal(ValidationType.List, workbook.Worksheets[0].Validations[0].Type);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithValidation("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithValidation("test_session_edit.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, validationIndex: 0, validationType: "WholeNumber", formula1: "0",
            formula2: "100");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(ValidationType.WholeNumber, workbook.Worksheets[0].Validations[0].Type);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithValidation("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        var beforeWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(beforeWorkbook.Worksheets[0].Validations.Count > 0);
        _tool.Execute("delete", sessionId: sessionId, validationIndex: 0);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].Validations);
    }

    [Fact]
    public void SetMessages_WithSessionId_ShouldSetInMemory()
    {
        var workbookPath = CreateWorkbookWithValidation("test_session_set_messages.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("set_messages", sessionId: sessionId, validationIndex: 0, inputMessage: "Session message");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("Session message", workbook.Worksheets[0].Validations[0].InputMessage);
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