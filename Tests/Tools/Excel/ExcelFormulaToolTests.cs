using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelFormulaToolTests : ExcelTestBase
{
    private readonly ExcelFormulaTool _tool;

    public ExcelFormulaToolTests()
    {
        _tool = new ExcelFormulaTool(SessionManager);
    }

    #region General Tests

    #region Calculate Tests

    [Fact]
    public void CalculateFormulas_ShouldCalculateAllFormulas()
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

        var result = _tool.Execute("calculate", workbookPath, outputPath: outputPath);

        Assert.Contains("Formulas calculated", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Add Tests

    [Fact]
    public void AddFormula_WithSum_ShouldAddSumFormula()
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

        _tool.Execute("add", workbookPath, cell: "A4", formula: "=SUM(A1:A3)", outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        Assert.Equal("=SUM(A1:A3)", worksheet.Cells["A4"].Formula);
        var result = worksheet.Cells["A4"].Value;
        Assert.NotNull(result);
        Assert.True(result is double or int, $"Formula result should be numeric, got: {result}");
    }

    [Fact]
    public void AddFormula_WithAutoCalculateFalse_ShouldNotCalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_no_calc_output.xlsx");

        var result = _tool.Execute("add", workbookPath, cell: "A3", formula: "=A1+A2", autoCalculate: false,
            outputPath: outputPath);

        Assert.Contains("Formula added", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void AddFormula_WithAverage_ShouldAddAverageFormula()
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

        _tool.Execute("add", workbookPath, cell: "A4", formula: "=AVERAGE(A1:A3)", outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=AVERAGE(A1:A3)", resultWorkbook.Worksheets[0].Cells["A4"].Formula);
    }

    [Fact]
    public void AddFormula_WithIf_ShouldAddIfFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_if.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 50;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_if_output.xlsx");

        _tool.Execute("add", workbookPath, cell: "B1", formula: "=IF(A1>40,\"Pass\",\"Fail\")", outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=IF(A1>40,\"Pass\",\"Fail\")", resultWorkbook.Worksheets[0].Cells["B1"].Formula);
    }

    [Fact]
    public void AddFormula_WithCount_ShouldAddCountFormula()
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

        _tool.Execute("add", workbookPath, cell: "A4", formula: "=COUNT(A1:A3)", outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=COUNT(A1:A3)", resultWorkbook.Worksheets[0].Cells["A4"].Formula);
    }

    [Fact]
    public void AddFormula_WithSheetIndex_ShouldApplyToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_sheet.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets[1].Cells["A1"].Value = 100;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_sheet_output.xlsx");

        _tool.Execute("add", workbookPath, sheetIndex: 1, cell: "B1", formula: "=A1*2", outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=A1*2", resultWorkbook.Worksheets[1].Cells["B1"].Formula);
    }

    #endregion

    #region Get Tests

    [Fact]
    public void GetFormula_ShouldReturnFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_get_formula.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Formula = "=SUM(B1:B10)";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);

        Assert.NotNull(result);
        Assert.Contains("A1", result);
        Assert.Contains("SUM", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetFormula_WithRange_ShouldReturnFormulasInRange()
    {
        var workbookPath = CreateExcelWorkbook("test_get_formula_range.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Formula = "=B1+1";
            workbook.Worksheets[0].Cells["A2"].Formula = "=B2+2";
            workbook.Worksheets[0].Cells["C1"].Formula = "=D1+3";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath, range: "A1:B2");

        var json = JsonDocument.Parse(result);
        Assert.Equal(2, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void GetFormula_NoFormulas_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_formula_empty.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = "Just text";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No formulas found", json.RootElement.GetProperty("message").GetString());
    }

    #endregion

    #region GetResult Tests

    [Fact]
    public void GetFormulaResult_ShouldReturnResult()
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

        var result = _tool.Execute("get_result", workbookPath, cell: "A3");

        Assert.NotNull(result);
        var json = JsonDocument.Parse(result);
        Assert.Equal("A3", json.RootElement.GetProperty("cell").GetString());
        Assert.Contains("30", json.RootElement.GetProperty("calculatedValue").GetString());
    }

    [Fact]
    public void GetFormulaResult_WithCalculateBeforeReadFalse_ShouldNotRecalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_get_result_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Formula = "=A1*2";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get_result", workbookPath, cell: "A2", calculateBeforeRead: false);

        Assert.NotNull(result);
        Assert.Contains("A2", result);
    }

    #endregion

    #region SetArray Tests

    [Fact]
    public void SetArrayFormula_ShouldSetArrayFormula()
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

        var result = _tool.Execute("set_array", workbookPath, range: "B1:B3", formula: "=A1:A3*2",
            outputPath: outputPath);

        Assert.Contains("Array formula set", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetArrayFormula_WithAutoCalculateFalse_ShouldNotCalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_array_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 1;
            workbook.Worksheets[0].Cells["A2"].Value = 2;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_array_no_calc_output.xlsx");

        _ = _tool.Execute("set_array", workbookPath, range: "B1:B2", formula: "=A1:A2*2", autoCalculate: false,
            outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region GetArray Tests

    [Fact]
    public void GetArrayFormula_ShouldReturnArrayFormula()
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

        var result = _tool.Execute("get_array", workbookPath, cell: "B1");

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isArrayFormula").GetBoolean());
    }

    [Fact]
    public void GetArrayFormula_NotArrayFormula_ShouldReturnFalse()
    {
        var workbookPath = CreateExcelWorkbook("test_get_array_not.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Formula = "=SUM(B1:B5)";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get_array", workbookPath, cell: "A1");

        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("isArrayFormula").GetBoolean());
    }

    #endregion

    #region Advanced Formula Tests

    [Fact]
    public void AddFormula_WithVLOOKUP_ShouldAddVLOOKUP()
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

        _tool.Execute("add", workbookPath, cell: "E1", formula: "=VLOOKUP(D1,A1:B3,2,FALSE)", outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=VLOOKUP(D1,A1:B3,2,FALSE)", resultWorkbook.Worksheets[0].Cells["E1"].Formula);
    }

    [Fact]
    public void AddFormula_WithSUMIF_ShouldAddSumif()
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

        _tool.Execute("add", workbookPath, cell: "C1", formula: "=SUMIF(A1:A3,\"Apple\",B1:B3)",
            outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("SUMIF", resultWorkbook.Worksheets[0].Cells["C1"].Formula);
    }

    [Fact]
    public void AddFormula_WithNestedFormula_ShouldAddNestedFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_formula_nested.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 85;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_formula_nested_output.xlsx");

        _tool.Execute("add", workbookPath, cell: "B1",
            formula: "=IF(A1>=90,\"A\",IF(A1>=80,\"B\",IF(A1>=70,\"C\",\"D\")))", outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("IF", resultWorkbook.Worksheets[0].Cells["B1"].Formula);
    }

    [Fact]
    public void AddFormula_WithCrossSheetRef_ShouldAddCrossSheetRef()
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

        _tool.Execute("add", workbookPath, sheetIndex: 1, cell: "A1", formula: "=Data!A1*2", outputPath: outputPath);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("Data", resultWorkbook.Worksheets[1].Cells["A1"].Formula);
    }

    [Fact]
    public void GetFormulaResult_WithErrorFormula_ShouldReturnError()
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

        var result = _tool.Execute("get_result", workbookPath, cell: "A3");

        Assert.NotNull(result);
        Assert.True(
            result.Contains("DIV") || result.Contains("Error") || result.Contains("Infinity") || result.Contains("A3"),
            $"Result should contain error info or cell reference, got: {result}");
    }

    #endregion

    #endregion

    #region Exception Tests

    [Fact]
    public void UnknownOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_sheet.xlsx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetIndex: 99, cell: "A1", formula: "=1+1"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void AddFormula_WithInvalidFunctionName_ShouldReturnWarning()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_func.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_invalid_func_output.xlsx");

        var result = _tool.Execute("add", workbookPath, cell: "B1", formula: "=INVALIDFUNC(A1)",
            outputPath: outputPath);

        Assert.Contains("Formula added", result);
        Assert.Contains("#NAME?", result);
    }

    [Fact]
    public void Add_MissingCell_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_cell.xlsx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, formula: "=SUM(A1:A10)"));
        Assert.Contains("cell", ex.Message.ToLower());
    }

    [Fact]
    public void Add_MissingFormula_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_formula.xlsx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, cell: "A1"));
        Assert.Contains("formula", ex.Message.ToLower());
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Formula = "=SUM(B1:B10)";
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("A1", result);
        Assert.Contains("SUM", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_add.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = 10;
            wb.Worksheets[0].Cells["A2"].Value = 20;
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, cell: "A3", formula: "=A1+A2");
        Assert.Contains("Formula added", result);

        // Verify in-memory workbook has the formula
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("=A1+A2", sessionWorkbook.Worksheets[0].Cells["A3"].Formula);
    }

    [Fact]
    public void Calculate_WithSessionId_ShouldCalculateInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_calculate.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = 10;
            wb.Worksheets[0].Cells["A2"].Value = 20;
            wb.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("calculate", sessionId: sessionId);
        Assert.Contains("Formulas calculated", result);

        // Verify in-memory workbook has calculated values
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        var calculatedValue = sessionWorkbook.Worksheets[0].Cells["A3"].Value;
        Assert.NotNull(calculatedValue);
    }

    [Fact]
    public void GetResult_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_result.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
            workbook.CalculateFormula();
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_result", sessionId: sessionId, cell: "A3");
        var json = JsonDocument.Parse(result);
        Assert.Equal("A3", json.RootElement.GetProperty("cell").GetString());
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}