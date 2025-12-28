using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelFormulaToolTests : ExcelTestBase
{
    private readonly ExcelFormulaTool _tool = new();

    #region Calculate Tests

    [Fact]
    public async Task CalculateFormulas_ShouldCalculateAllFormulas()
    {
        var workbookPath = CreateExcelWorkbook("test_calculate_formulas.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_calculate_formulas_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "calculate",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Formulas calculated", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Add Tests

    [Fact]
    public async Task AddFormula_WithSum_ShouldAddSumFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_sum.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Value = 30;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_sum_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A4",
            ["formula"] = "=SUM(A1:A3)"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("=SUM(A1:A3)", worksheet.Cells["A4"].Formula);
        var result = worksheet.Cells["A4"].Value;
        Assert.NotNull(result);
        Assert.True(result is double or int, $"Formula result should be numeric, got: {result}");
    }

    [Fact]
    public async Task AddFormula_WithAutoCalculateFalse_ShouldNotCalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_no_calc_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A3",
            ["formula"] = "=A1+A2",
            ["autoCalculate"] = false
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Formula added", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task AddFormula_WithAverage_ShouldAddAverageFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_average.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Value = 30;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_average_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A4",
            ["formula"] = "=AVERAGE(A1:A3)"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=AVERAGE(A1:A3)", resultWorkbook.Worksheets[0].Cells["A4"].Formula);
    }

    [Fact]
    public async Task AddFormula_WithIf_ShouldAddIfFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_if.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 50;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_if_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "B1",
            ["formula"] = "=IF(A1>40,\"Pass\",\"Fail\")"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=IF(A1>40,\"Pass\",\"Fail\")", resultWorkbook.Worksheets[0].Cells["B1"].Formula);
    }

    [Fact]
    public async Task AddFormula_WithCount_ShouldAddCountFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_count.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Value = "";
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_count_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A4",
            ["formula"] = "=COUNT(A1:A3)"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=COUNT(A1:A3)", resultWorkbook.Worksheets[0].Cells["A4"].Formula);
    }

    [Fact]
    public async Task AddFormula_WithSheetIndex_ShouldApplyToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_sheet.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets[1].Cells["A1"].Value = 100;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_sheet_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 1,
            ["cell"] = "B1",
            ["formula"] = "=A1*2"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=A1*2", resultWorkbook.Worksheets[1].Cells["B1"].Formula);
    }

    #endregion

    #region Get Tests

    [Fact]
    public async Task GetFormula_ShouldReturnFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_get_formula.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Formula = "=SUM(B1:B10)";
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.NotNull(result);
        Assert.Contains("A1", result);
        Assert.Contains("SUM", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetFormula_WithRange_ShouldReturnFormulasInRange()
    {
        var workbookPath = CreateExcelWorkbook("test_get_formula_range.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Formula = "=B1+1";
            workbook.Worksheets[0].Cells["A2"].Formula = "=B2+2";
            workbook.Worksheets[0].Cells["C1"].Formula = "=D1+3";
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["range"] = "A1:B2"
        };

        var result = await _tool.ExecuteAsync(arguments);

        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public async Task GetFormula_NoFormulas_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_formula_empty.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Just text";
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        var result = await _tool.ExecuteAsync(arguments);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No formulas found", json.RootElement.GetProperty("message").GetString());
    }

    #endregion

    #region GetResult Tests

    [Fact]
    public async Task GetFormulaResult_ShouldReturnResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_formula_result.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
            workbook.CalculateFormula();
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_result",
            ["path"] = workbookPath,
            ["cell"] = "A3"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.Equal("A3", json.RootElement.GetProperty("cell").GetString());
        Assert.Contains("30", json.RootElement.GetProperty("calculatedValue").GetString());
    }

    [Fact]
    public async Task GetFormulaResult_WithCalculateBeforeReadFalse_ShouldNotRecalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_get_result_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Formula = "=A1*2";
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_result",
            ["path"] = workbookPath,
            ["cell"] = "A2",
            ["calculateBeforeRead"] = false
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.NotNull(result);
        Assert.Contains("A2", result);
    }

    #endregion

    #region SetArray Tests

    [Fact]
    public async Task SetArrayFormula_ShouldSetArrayFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_array_formula.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 1;
            workbook.Worksheets[0].Cells["A2"].Value = 2;
            workbook.Worksheets[0].Cells["A3"].Value = 3;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_array_formula_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_array",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "B1:B3",
            ["formula"] = "=A1:A3*2"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Array formula set", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetArrayFormula_WithAutoCalculateFalse_ShouldNotCalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_array_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 1;
            workbook.Worksheets[0].Cells["A2"].Value = 2;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_array_no_calc_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_array",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "B1:B2",
            ["formula"] = "=A1:A2*2",
            ["autoCalculate"] = false
        };

        _ = await _tool.ExecuteAsync(arguments);

        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region GetArray Tests

    [Fact]
    public async Task GetArrayFormula_ShouldReturnArrayFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_get_array_formula.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 1;
            workbook.Worksheets[0].Cells["A2"].Value = 2;
#pragma warning disable CS0618
            workbook.Worksheets[0].Cells["B1"].SetArrayFormula("=A1:A2*2", 2, 1);
#pragma warning restore CS0618
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_array",
            ["path"] = workbookPath,
            ["cell"] = "B1"
        };

        var result = await _tool.ExecuteAsync(arguments);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isArrayFormula").GetBoolean());
    }

    [Fact]
    public async Task GetArrayFormula_NotArrayFormula_ShouldReturnFalse()
    {
        var workbookPath = CreateExcelWorkbook("test_get_array_not.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Formula = "=SUM(B1:B5)";
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_array",
            ["path"] = workbookPath,
            ["cell"] = "A1"
        };

        var result = await _tool.ExecuteAsync(arguments);

        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("isArrayFormula").GetBoolean());
    }

    #endregion

    #region Advanced Formula Tests

    [Fact]
    public async Task AddFormula_WithVLOOKUP_ShouldAddVLOOKUP()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_vlookup.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].Value = "ID001";
            worksheet.Cells["B1"].Value = "Apple";
            worksheet.Cells["A2"].Value = "ID002";
            worksheet.Cells["B2"].Value = "Banana";
            worksheet.Cells["A3"].Value = "ID003";
            worksheet.Cells["B3"].Value = "Cherry";
            worksheet.Cells["D1"].Value = "ID002";
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_vlookup_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "E1",
            ["formula"] = "=VLOOKUP(D1,A1:B3,2,FALSE)"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=VLOOKUP(D1,A1:B3,2,FALSE)", resultWorkbook.Worksheets[0].Cells["E1"].Formula);
    }

    [Fact]
    public async Task AddFormula_WithSUMIF_ShouldAddSumif()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_sumif.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].Value = "Apple";
            worksheet.Cells["B1"].Value = 10;
            worksheet.Cells["A2"].Value = "Banana";
            worksheet.Cells["B2"].Value = 20;
            worksheet.Cells["A3"].Value = "Apple";
            worksheet.Cells["B3"].Value = 30;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_sumif_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "C1",
            ["formula"] = "=SUMIF(A1:A3,\"Apple\",B1:B3)"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("SUMIF", resultWorkbook.Worksheets[0].Cells["C1"].Formula);
    }

    [Fact]
    public async Task AddFormula_WithNestedFormula_ShouldAddNestedFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_nested.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 85;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_nested_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "B1",
            ["formula"] = "=IF(A1>=90,\"A\",IF(A1>=80,\"B\",IF(A1>=70,\"C\",\"D\")))"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("IF", resultWorkbook.Worksheets[0].Cells["B1"].Formula);
    }

    [Fact]
    public async Task AddFormula_WithCrossSheetRef_ShouldAddCrossSheetRef()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_cross_sheet.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Data";
            sheet1.Cells["A1"].Value = 100;
            workbook.Worksheets.Add("Summary");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_cross_sheet_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "A1",
            ["sheetIndex"] = 1,
            ["formula"] = "=Data!A1*2"
        };

        await _tool.ExecuteAsync(arguments);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("Data", resultWorkbook.Worksheets[1].Cells["A1"].Formula);
    }

    [Fact]
    public async Task GetFormulaResult_WithErrorFormula_ShouldReturnError()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_error.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 0;
            workbook.Worksheets[0].Cells["A3"].Formula = "=A1/A2";
            workbook.CalculateFormula();
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get_result",
            ["path"] = workbookPath,
            ["cell"] = "A3"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.NotNull(result);
        Assert.True(
            result.Contains("DIV") || result.Contains("Error") || result.Contains("Infinity") || result.Contains("A3"),
            $"Result should contain error info or cell reference, got: {result}");
    }

    #endregion

    #region Error Handling Tests

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
    public async Task InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["cell"] = "A1",
            ["formula"] = "=1+1"
        };

        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public async Task AddFormula_WithInvalidFunctionName_ShouldReturnWarning()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_func.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_invalid_func_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["cell"] = "B1",
            ["formula"] = "=INVALIDFUNC(A1)"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Formula added", result);
        Assert.Contains("#NAME?", result);
    }

    #endregion
}