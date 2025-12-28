using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelDataValidationToolTests : ExcelTestBase
{
    private readonly ExcelDataValidationTool _tool = new();

    [Fact]
    public async Task AddDataValidation_WithList_ShouldAddListValidation()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_validation.xlsx");
        var outputPath = CreateTestFilePath("test_add_validation_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A10",
            ["validationType"] = "List",
            ["formula1"] = "Option1,Option2,Option3"
        };

        var result = await _tool.ExecuteAsync(arguments);

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
    public async Task AddDataValidation_WithList_InCellDropDownFalse_ShouldDisableDropdown()
    {
        var workbookPath = CreateExcelWorkbook("test_add_list_no_dropdown.xlsx");
        var outputPath = CreateTestFilePath("test_add_list_no_dropdown_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A10",
            ["validationType"] = "List",
            ["formula1"] = "1,2,3",
            ["inCellDropDown"] = false
        };

        await _tool.ExecuteAsync(arguments);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal(ValidationType.List, validation.Type);
        Assert.False(validation.InCellDropDown);
    }

    [Fact]
    public async Task AddDataValidation_WithWholeNumber_ShouldAddNumberValidation()
    {
        var workbookPath = CreateExcelWorkbook("test_add_number_validation.xlsx");
        var outputPath = CreateTestFilePath("test_add_number_validation_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A10",
            ["validationType"] = "WholeNumber",
            ["formula1"] = "0",
            ["formula2"] = "100"
        };

        await _tool.ExecuteAsync(arguments);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal(ValidationType.WholeNumber, validation.Type);
        Assert.Equal(OperatorType.Between, validation.Operator);
    }

    [Fact]
    public async Task AddDataValidation_WithOperatorGreaterThan_ShouldUseGreaterThanOperator()
    {
        var workbookPath = CreateExcelWorkbook("test_add_greater_than.xlsx");
        var outputPath = CreateTestFilePath("test_add_greater_than_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A10",
            ["validationType"] = "WholeNumber",
            ["operatorType"] = "GreaterThan",
            ["formula1"] = "0"
        };

        await _tool.ExecuteAsync(arguments);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal(ValidationType.WholeNumber, validation.Type);
        Assert.Equal(OperatorType.GreaterThan, validation.Operator);
    }

    [Fact]
    public async Task AddDataValidation_WithOperatorLessThan_ShouldUseLessThanOperator()
    {
        var workbookPath = CreateExcelWorkbook("test_add_less_than.xlsx");
        var outputPath = CreateTestFilePath("test_add_less_than_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "B1:B10",
            ["validationType"] = "Decimal",
            ["operatorType"] = "LessThan",
            ["formula1"] = "100.5"
        };

        await _tool.ExecuteAsync(arguments);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal(ValidationType.Decimal, validation.Type);
        Assert.Equal(OperatorType.LessThan, validation.Operator);
    }

    [Fact]
    public async Task AddDataValidation_WithMessages_ShouldSetMessages()
    {
        var workbookPath = CreateExcelWorkbook("test_add_with_messages.xlsx");
        var outputPath = CreateTestFilePath("test_add_with_messages_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:A10",
            ["validationType"] = "List",
            ["formula1"] = "Yes,No",
            ["inputMessage"] = "Please select Yes or No",
            ["errorMessage"] = "Invalid selection"
        };

        await _tool.ExecuteAsync(arguments);

        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.Equal("Please select Yes or No", validation.InputMessage);
        Assert.Equal("Invalid selection", validation.ErrorMessage);
        Assert.True(validation.ShowInput);
        Assert.True(validation.ShowError);
    }

    [Fact]
    public async Task GetDataValidation_ShouldReturnValidationInfo()
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

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        Assert.True(json.RootElement.GetProperty("items").GetArrayLength() > 0);
    }

    [Fact]
    public async Task GetDataValidation_EmptyWorksheet_ShouldReturnEmptyList()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Equal(0, json.RootElement.GetProperty("items").GetArrayLength());
    }

    [Fact]
    public async Task SetMessages_ShouldSetInputAndErrorMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_set_messages.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        worksheet.Validations.Add(area);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_set_messages_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_messages",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["validationIndex"] = 0,
            ["inputMessage"] = "Please select a value",
            ["errorMessage"] = "Invalid value selected"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Updated data validation", result);
        Assert.Contains("InputMessage=Please select a value", result);
        Assert.Contains("ErrorMessage=Invalid value selected", result);

        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var validation = resultWorksheet.Validations[0];
        Assert.Equal("Please select a value", validation.InputMessage);
        Assert.Equal("Invalid value selected", validation.ErrorMessage);
    }

    [Fact]
    public async Task SetMessages_ClearMessage_ShouldClearAndDisableShow()
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
        var arguments = new JsonObject
        {
            ["operation"] = "set_messages",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["validationIndex"] = 0,
            ["inputMessage"] = ""
        };

        await _tool.ExecuteAsync(arguments);

        var resultWorkbook = new Workbook(outputPath);
        var validation = resultWorkbook.Worksheets[0].Validations[0];
        Assert.True(string.IsNullOrEmpty(validation.InputMessage));
        Assert.False(validation.ShowInput);
    }

    [Fact]
    public async Task DeleteDataValidation_ShouldDeleteValidation()
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
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["validationIndex"] = 0
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Deleted data validation #0", result);
        Assert.Contains("remaining: 0", result);

        var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].Validations);
    }

    [Fact]
    public async Task DeleteDataValidation_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_delete_invalid.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["validationIndex"] = 99
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public async Task EditDataValidation_ShouldEditValidation()
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["validationIndex"] = 0,
            ["validationType"] = "WholeNumber",
            ["formula1"] = "0",
            ["formula2"] = "100"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Edited data validation #0", result);
        Assert.Contains("Type=WholeNumber", result);

        var resultWorkbook = new Workbook(outputPath);
        var resultValidation = resultWorkbook.Worksheets[0].Validations[0];
        Assert.Equal(ValidationType.WholeNumber, resultValidation.Type);
    }

    [Fact]
    public async Task EditDataValidation_ChangeOperatorType_ShouldUpdateOperator()
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["validationIndex"] = 0,
            ["operatorType"] = "GreaterOrEqual"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Operator=GreaterOrEqual", result);

        var resultWorkbook = new Workbook(outputPath);
        var validation = resultWorkbook.Worksheets[0].Validations[0];
        Assert.Equal(OperatorType.GreaterOrEqual, validation.Operator);
    }

    [Fact]
    public async Task EditDataValidation_ChangeInCellDropDown_ShouldUpdateDropdown()
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["validationIndex"] = 0,
            ["inCellDropDown"] = false
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("InCellDropDown=False", result);

        var resultWorkbook = new Workbook(outputPath);
        var validation = resultWorkbook.Worksheets[0].Validations[0];
        Assert.False(validation.InCellDropDown);
    }

    [Fact]
    public async Task EditDataValidation_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_edit_invalid.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["validationIndex"] = 99,
            ["formula1"] = "test"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public async Task SetMessages_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_set_messages_invalid.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_messages",
            ["path"] = workbookPath,
            ["validationIndex"] = 99,
            ["inputMessage"] = "test"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public async Task AddDataValidation_InvalidValidationType_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_type.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["range"] = "A1:A10",
            ["validationType"] = "InvalidType",
            ["formula1"] = "test"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unsupported validation type", ex.Message);
    }

    [Fact]
    public async Task AddDataValidation_InvalidOperatorType_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_operator.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["range"] = "A1:A10",
            ["validationType"] = "WholeNumber",
            ["operatorType"] = "InvalidOperator",
            ["formula1"] = "0"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unsupported operator type", ex.Message);
    }

    [Fact]
    public async Task UnknownOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = workbookPath
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public async Task AddDataValidation_AllValidationTypes_ShouldWork()
    {
        var validationTypes = new[] { "WholeNumber", "Decimal", "Date", "Time", "TextLength", "Custom" };

        foreach (var validationType in validationTypes)
        {
            var workbookPath = CreateExcelWorkbook($"test_type_{validationType}.xlsx");
            var outputPath = CreateTestFilePath($"test_type_{validationType}_output.xlsx");
            var arguments = new JsonObject
            {
                ["operation"] = "add",
                ["path"] = workbookPath,
                ["outputPath"] = outputPath,
                ["range"] = "A1:A10",
                ["validationType"] = validationType,
                ["formula1"] = "1"
            };

            var result = await _tool.ExecuteAsync(arguments);
            Assert.Contains($"type: {validationType}", result);
        }
    }

    [Fact]
    public async Task AddDataValidation_AllOperatorTypes_ShouldWork()
    {
        var operatorTypes = new[]
            { "Between", "Equal", "NotEqual", "GreaterThan", "LessThan", "GreaterOrEqual", "LessOrEqual" };

        foreach (var operatorType in operatorTypes)
        {
            var workbookPath = CreateExcelWorkbook($"test_op_{operatorType}.xlsx");
            var outputPath = CreateTestFilePath($"test_op_{operatorType}_output.xlsx");
            var arguments = new JsonObject
            {
                ["operation"] = "add",
                ["path"] = workbookPath,
                ["outputPath"] = outputPath,
                ["range"] = "A1:A10",
                ["validationType"] = "WholeNumber",
                ["operatorType"] = operatorType,
                ["formula1"] = "0",
                ["formula2"] = operatorType == "Between" ? "100" : null
            };

            var result = await _tool.ExecuteAsync(arguments);
            Assert.Contains("Data validation added", result);
        }
    }
}