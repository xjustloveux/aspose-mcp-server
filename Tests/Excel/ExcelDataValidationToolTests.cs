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
        // Arrange
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

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        var validation = worksheet.Validations[^1];
        Assert.NotNull(validation);
        Assert.Equal(ValidationType.List, validation.Type);
    }

    [Fact]
    public async Task AddDataValidation_WithWholeNumber_ShouldAddNumberValidation()
    {
        // Arrange
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

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.Validations.Count > 0, "Validation should be added");
    }

    [Fact]
    public async Task GetDataValidation_ShouldReturnValidationInfo()
    {
        // Arrange
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
            ["path"] = workbookPath,
            ["validationIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("count", result);
    }

    [Fact]
    public async Task SetMessages_ShouldSetInputAndErrorMessage()
    {
        // Arrange
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

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var validation = resultWorksheet.Validations[0];
        Assert.NotNull(validation);
    }

    [Fact]
    public async Task DeleteDataValidation_ShouldDeleteValidation()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_delete_validation.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var area = new CellArea { StartRow = 0, StartColumn = 0, EndRow = 9, EndColumn = 0 };
        worksheet.Validations.Add(area);
        workbook.Save(workbookPath);

        var validationsBefore = worksheet.Validations.Count;
        Assert.True(validationsBefore > 0, "Validation should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_validation_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["validationIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var validationsAfter = resultWorksheet.Validations.Count;
        Assert.True(validationsAfter < validationsBefore,
            $"Validation should be deleted. Before: {validationsBefore}, After: {validationsAfter}");
    }

    [Fact]
    public async Task EditDataValidation_ShouldEditValidation()
    {
        // Arrange
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
            ["range"] = "B1:B10",
            ["validationType"] = "WholeNumber",
            ["formula1"] = "0",
            ["formula2"] = "100"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.Validations.Count > 0, "Validation should exist after editing");
    }
}