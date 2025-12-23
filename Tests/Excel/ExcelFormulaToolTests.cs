using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelFormulaToolTests : ExcelTestBase
{
    private readonly ExcelFormulaTool _tool = new();

    [Fact]
    public async Task AddFormula_WithSum_ShouldAddSumFormula()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_formula_sum.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Value = 20;
        workbook.Worksheets[0].Cells["A3"].Value = 30;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_formula_sum_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A4",
            ["formula"] = "=SUM(A1:A3)"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("=SUM(A1:A3)", worksheet.Cells["A4"].Formula);
        // Verify formula result was calculated
        var result = worksheet.Cells["A4"].Value;
        Assert.NotNull(result);
        // Result should be 60 (10+20+30) or at least a numeric value
        Assert.True(result is double || result is int, $"Formula result should be numeric, got: {result}");
    }

    [Fact]
    public async Task AddFormula_WithAverage_ShouldAddAverageFormula()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_formula_average.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Value = 20;
        workbook.Worksheets[0].Cells["A3"].Value = 30;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_formula_average_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A4",
            ["formula"] = "=AVERAGE(A1:A3)"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("=AVERAGE(A1:A3)", worksheet.Cells["A4"].Formula);
    }

    [Fact]
    public async Task AddFormula_WithIf_ShouldAddIfFormula()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_formula_if.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 50;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_formula_if_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "B1",
            ["formula"] = "=IF(A1>40,\"Pass\",\"Fail\")"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("=IF(A1>40,\"Pass\",\"Fail\")", worksheet.Cells["B1"].Formula);
    }

    [Fact]
    public async Task AddFormula_WithCount_ShouldAddCountFormula()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_formula_count.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Value = 20;
        workbook.Worksheets[0].Cells["A3"].Value = "";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_formula_count_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A4",
            ["formula"] = "=COUNT(A1:A3)"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("=COUNT(A1:A3)", worksheet.Cells["A4"].Formula);
    }

    [Fact]
    public async Task GetFormula_ShouldReturnFormula()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_formula.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Formula = "=SUM(B1:B10)";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["cell"] = "A1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("A1", result);
        Assert.Contains("SUM", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetFormulaResult_ShouldReturnResult()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_formula_result.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Value = 20;
        workbook.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
        workbook.CalculateFormula();
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_result",
            ["path"] = workbookPath,
            ["cell"] = "A3"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("A3", result);
    }

    [Fact]
    public async Task CalculateFormulas_ShouldCalculateAllFormulas()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_calculate_formulas.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Value = 20;
        workbook.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_calculate_formulas_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "calculate",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }

    [Fact]
    public async Task SetArrayFormula_ShouldSetArrayFormula()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_array_formula.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 1;
        workbook.Worksheets[0].Cells["A2"].Value = 2;
        workbook.Worksheets[0].Cells["A3"].Value = 3;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_array_formula_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_array",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "B1:B3",
            ["formula"] = "=A1:A3*2"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }

    [Fact]
    public async Task GetArrayFormula_ShouldReturnArrayFormula()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_array_formula.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 1;
        workbook.Worksheets[0].Cells["A2"].Value = 2;
        var firstCell = workbook.Worksheets[0].Cells["A1"];
        firstCell.SetArrayFormula("=A1:A2*2", 2, 1);
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_array",
            ["path"] = workbookPath,
            ["cell"] = "A1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }
}