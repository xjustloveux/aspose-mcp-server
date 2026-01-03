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

    #region General Tests

    [Fact]
    public void AddDataValidation_WithList_ShouldAddListValidation()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_validation.xlsx");
        var outputPath = CreateTestFilePath("test_add_validation_output.xlsx");

        var result = _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A10",
            validationType: "List",
            formula1: "Option1,Option2,Option3",
            outputPath: outputPath);

        Assert.Contains("Data validation added", result);
        Assert.Contains("index: 0", result);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.NotNull(validation);
        Assert.Equal(ValidationType.List, validation.Type);
        Assert.True(validation.InCellDropDown);
    }

    [Fact]
    public void AddDataValidation_WithList_InCellDropDownFalse_ShouldDisableDropdown()
    {
        var workbookPath = CreateExcelWorkbook("test_add_list_no_dropdown.xlsx");
        var outputPath = CreateTestFilePath("test_add_list_no_dropdown_output.xlsx");

        _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A10",
            validationType: "List",
            formula1: "1,2,3",
            inCellDropDown: false,
            outputPath: outputPath);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal(ValidationType.List, validation.Type);
        Assert.False(validation.InCellDropDown);
    }

    [Fact]
    public void AddDataValidation_WithWholeNumber_ShouldAddNumberValidation()
    {
        var workbookPath = CreateExcelWorkbook("test_add_number_validation.xlsx");
        var outputPath = CreateTestFilePath("test_add_number_validation_output.xlsx");

        _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A10",
            validationType: "WholeNumber",
            formula1: "0",
            formula2: "100",
            outputPath: outputPath);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal(ValidationType.WholeNumber, validation.Type);
        Assert.Equal(OperatorType.Between, validation.Operator);
    }

    [Fact]
    public void AddDataValidation_WithOperatorGreaterThan_ShouldUseGreaterThanOperator()
    {
        var workbookPath = CreateExcelWorkbook("test_add_greater_than.xlsx");
        var outputPath = CreateTestFilePath("test_add_greater_than_output.xlsx");

        _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A10",
            validationType: "WholeNumber",
            operatorType: "GreaterThan",
            formula1: "0",
            outputPath: outputPath);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal(ValidationType.WholeNumber, validation.Type);
        Assert.Equal(OperatorType.GreaterThan, validation.Operator);
    }

    [Fact]
    public void AddDataValidation_WithOperatorLessThan_ShouldUseLessThanOperator()
    {
        var workbookPath = CreateExcelWorkbook("test_add_less_than.xlsx");
        var outputPath = CreateTestFilePath("test_add_less_than_output.xlsx");

        _tool.Execute(
            "add",
            workbookPath,
            range: "B1:B10",
            validationType: "Decimal",
            operatorType: "LessThan",
            formula1: "100.5",
            outputPath: outputPath);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal(ValidationType.Decimal, validation.Type);
        Assert.Equal(OperatorType.LessThan, validation.Operator);
    }

    [Fact]
    public void AddDataValidation_WithMessages_ShouldSetMessages()
    {
        var workbookPath = CreateExcelWorkbook("test_add_with_messages.xlsx");
        var outputPath = CreateTestFilePath("test_add_with_messages_output.xlsx");

        _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A10",
            validationType: "List",
            formula1: "Yes,No",
            inputMessage: "Please select Yes or No",
            errorMessage: "Invalid selection",
            outputPath: outputPath);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal("Please select Yes or No", validation.InputMessage);
        Assert.Equal("Invalid selection", validation.ErrorMessage);
        Assert.True(validation.ShowInput);
        Assert.True(validation.ShowError);
    }

    [Fact]
    public void GetDataValidation_ShouldReturnValidationInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_validation.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var validationIndex = worksheet.Validations.Add(area);
        var validation = worksheet.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.Formula1 = "1,2,3";
        workbook.Save(workbookPath);

        var result = _tool.Execute(
            "get",
            workbookPath);

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.True(json.RootElement.GetProperty("items").GetArrayLength() > 0);
    }

    [Fact]
    public void GetDataValidation_EmptyWorksheet_ShouldReturnEmptyList()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");

        var result = _tool.Execute(
            "get",
            workbookPath);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(0, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public void SetMessages_ShouldSetInputAndErrorMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_set_messages.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        worksheet.Validations.Add(area);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_set_messages_output.xlsx");

        var result = _tool.Execute(
            "set_messages",
            workbookPath,
            validationIndex: 0,
            inputMessage: "Please select a value",
            errorMessage: "Invalid value selected",
            outputPath: outputPath);

        Assert.Contains("Updated data validation", result);
        Assert.Contains("InputMessage=Please select a value", result);
        Assert.Contains("ErrorMessage=Invalid value selected", result);

        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var validationResult = resultWorksheet.Validations[0];
        Assert.Equal("Please select a value", validationResult.InputMessage);
        Assert.Equal("Invalid value selected", validationResult.ErrorMessage);
    }

    [Fact]
    public void SetMessages_ClearMessage_ShouldClearAndDisableShow()
    {
        var workbookPath = CreateExcelWorkbook("test_clear_messages.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var idx = worksheet.Validations.Add(area);
        worksheet.Validations[idx].InputMessage = "Old message";
        worksheet.Validations[idx].ShowInput = true;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_clear_messages_output.xlsx");

        _tool.Execute(
            "set_messages",
            workbookPath,
            validationIndex: 0,
            inputMessage: "",
            outputPath: outputPath);

        var resultWorkbook = new Workbook(outputPath);
        var validation = resultWorkbook.Worksheets[0].Validations[0];
        Assert.True(string.IsNullOrEmpty(validation.InputMessage));
        Assert.False(validation.ShowInput);
    }

    [Fact]
    public void DeleteDataValidation_ShouldDeleteValidation()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_validation.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        worksheet.Validations.Add(area);
        workbook.Save(workbookPath);

        var validationsBefore = worksheet.Validations.Count;
        Assert.True(validationsBefore > 0);

        var outputPath = CreateTestFilePath("test_delete_validation_output.xlsx");

        var result = _tool.Execute(
            "delete",
            workbookPath,
            validationIndex: 0,
            outputPath: outputPath);

        Assert.Contains("Deleted data validation #0", result);
        Assert.Contains("remaining: 0", result);

        var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].Validations);
    }

    [Fact]
    public void DeleteDataValidation_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            workbookPath,
            validationIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void EditDataValidation_ShouldEditValidation()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_validation.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var validationIndex = worksheet.Validations.Add(area);
        var validation = worksheet.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.Formula1 = "1,2,3";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_validation_output.xlsx");

        var result = _tool.Execute(
            "edit",
            workbookPath,
            validationIndex: 0,
            validationType: "WholeNumber",
            formula1: "0",
            formula2: "100",
            outputPath: outputPath);

        Assert.Contains("Edited data validation #0", result);
        Assert.Contains("Type=WholeNumber", result);

        var resultWorkbook = new Workbook(outputPath);
        var resultValidation = resultWorkbook.Worksheets[0].Validations[0];
        Assert.Equal(ValidationType.WholeNumber, resultValidation.Type);
    }

    [Fact]
    public void EditDataValidation_ChangeOperatorType_ShouldUpdateOperator()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_operator.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var idx = worksheet.Validations.Add(area);
        worksheet.Validations[idx].Type = ValidationType.WholeNumber;
        worksheet.Validations[idx].Formula1 = "0";
        worksheet.Validations[idx].Operator = OperatorType.Equal;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_operator_output.xlsx");

        var result = _tool.Execute(
            "edit",
            workbookPath,
            validationIndex: 0,
            operatorType: "GreaterOrEqual",
            outputPath: outputPath);

        Assert.Contains("Operator=GreaterOrEqual", result);

        var resultWorkbook = new Workbook(outputPath);
        var validation = resultWorkbook.Worksheets[0].Validations[0];
        Assert.Equal(OperatorType.GreaterOrEqual, validation.Operator);
    }

    [Fact]
    public void EditDataValidation_ChangeInCellDropDown_ShouldUpdateDropdown()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_dropdown.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var idx = worksheet.Validations.Add(area);
        worksheet.Validations[idx].Type = ValidationType.List;
        worksheet.Validations[idx].Formula1 = "A,B,C";
        worksheet.Validations[idx].InCellDropDown = true;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_dropdown_output.xlsx");

        var result = _tool.Execute(
            "edit",
            workbookPath,
            validationIndex: 0,
            inCellDropDown: false,
            outputPath: outputPath);

        Assert.Contains("InCellDropDown=False", result);

        var resultWorkbook = new Workbook(outputPath);
        var validation = resultWorkbook.Worksheets[0].Validations[0];
        Assert.False(validation.InCellDropDown);
    }

    [Fact]
    public void EditDataValidation_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_invalid.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            workbookPath,
            validationIndex: 99,
            formula1: "test"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void SetMessages_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_set_messages_invalid.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_messages",
            workbookPath,
            validationIndex: 99,
            inputMessage: "test"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void AddDataValidation_InvalidValidationType_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_type.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A10",
            validationType: "InvalidType",
            formula1: "test"));
        Assert.Contains("Unsupported validation type", ex.Message);
    }

    [Fact]
    public void AddDataValidation_InvalidOperatorType_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_operator.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            range: "A1:A10",
            validationType: "WholeNumber",
            operatorType: "InvalidOperator",
            formula1: "0"));
        Assert.Contains("Unsupported operator type", ex.Message);
    }

    [Fact]
    public void UnknownOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "unknown",
            workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddDataValidation_AllValidationTypes_ShouldWork()
    {
        var validationTypes = new[] { "WholeNumber", "Decimal", "Date", "Time", "TextLength", "Custom" };

        foreach (var validationType in validationTypes)
        {
            var workbookPath = CreateExcelWorkbook($"test_type_{validationType}.xlsx");
            var outputPath = CreateTestFilePath($"test_type_{validationType}_output.xlsx");

            var result = _tool.Execute(
                "add",
                workbookPath,
                range: "A1:A10",
                validationType: validationType,
                formula1: "1",
                outputPath: outputPath);
            Assert.Contains($"type: {validationType}", result);
        }
    }

    [Fact]
    public void AddDataValidation_AllOperatorTypes_ShouldWork()
    {
        var operatorTypes = new[]
            { "Between", "Equal", "NotEqual", "GreaterThan", "LessThan", "GreaterOrEqual", "LessOrEqual" };

        foreach (var operatorType in operatorTypes)
        {
            var workbookPath = CreateExcelWorkbook($"test_op_{operatorType}.xlsx");
            var outputPath = CreateTestFilePath($"test_op_{operatorType}_output.xlsx");

            var result = _tool.Execute(
                "add",
                workbookPath,
                range: "A1:A10",
                validationType: "WholeNumber",
                operatorType: operatorType,
                formula1: "0",
                formula2: operatorType == "Between" ? "100" : null,
                outputPath: outputPath);
            Assert.Contains("Data validation added", result);
        }
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_ex_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", workbookPath));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_range.xlsx");
        var outputPath = CreateTestFilePath("test_missing_range_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, validationType: "List", formula1: "A,B,C", outputPath: outputPath));

        Assert.Contains("range is required", ex.Message);
    }

    [Fact]
    public void Add_WithMissingValidationType_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_missing_type.xlsx");
        var outputPath = CreateTestFilePath("test_missing_type_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, range: "A1:A10", formula1: "A,B,C", outputPath: outputPath));

        Assert.Contains("validationType is required", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_validation.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var validationIndex = worksheet.Validations.Add(area);
        var validation = worksheet.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.Formula1 = "Session1,Session2,Session3";
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add_validation.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, range: "A1:A10", validationType: "List",
            formula1: "SessionOpt1,SessionOpt2");
        Assert.Contains("Data validation added", result);

        // Verify in-memory workbook has the validation
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(sessionWorkbook.Worksheets[0].Validations.Count > 0, "Validation should be added in memory");
        Assert.Equal(ValidationType.List, sessionWorkbook.Worksheets[0].Validations[0].Type);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_edit_validation.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        var validationIndex = worksheet.Validations.Add(area);
        var validation = worksheet.Validations[validationIndex];
        validation.Type = ValidationType.List;
        validation.Formula1 = "1,2,3";
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("edit", sessionId: sessionId, validationIndex: 0, validationType: "WholeNumber", formula1: "0",
            formula2: "100");

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(ValidationType.WholeNumber, sessionWorkbook.Worksheets[0].Validations[0].Type);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_delete_validation.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        worksheet.Validations.Add(area);
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);

        // Verify validation exists before delete
        var beforeWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(beforeWorkbook.Worksheets[0].Validations.Count > 0);
        _tool.Execute("delete", sessionId: sessionId, validationIndex: 0);

        // Assert - verify in-memory change
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(sessionWorkbook.Worksheets[0].Validations);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}