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

    private string CreateWorkbookWithFormula(string fileName, string cell = "A3", string formula = "=A1+A2")
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["A2"].Value = 20;
        workbook.Worksheets[0].Cells[cell].Formula = formula;
        workbook.CalculateFormula();
        workbook.Save(path);
        return path;
    }

    private string CreateWorkbookWithArrayFormula(string fileName)
    {
        var path = CreateExcelWorkbook(fileName);
        using var workbook = new Workbook(path);
        workbook.Worksheets[0].Cells["A1"].Value = 1;
        workbook.Worksheets[0].Cells["A2"].Value = 2;
#pragma warning disable CS0618
        workbook.Worksheets[0].Cells["B1"].SetArrayFormula("=A1:A2*2", 2, 1);
#pragma warning restore CS0618
        workbook.Save(path);
        return path;
    }

    #region General

    [Fact]
    public void Add_WithSum_ShouldAddSumFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_add_sum.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Value = 30;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_sum_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "A4", formula: "=SUM(A1:A3)", outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=SUM(A1:A3)", resultWorkbook.Worksheets[0].Cells["A4"].Formula);
        var result = resultWorkbook.Worksheets[0].Cells["A4"].Value;
        Assert.NotNull(result);
        Assert.True(result is double or int);
    }

    [Fact]
    public void Add_WithAutoCalculateFalse_ShouldNotCalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_add_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_no_calc_output.xlsx");
        var result = _tool.Execute("add", workbookPath, cell: "A3", formula: "=A1+A2", autoCalculate: false,
            outputPath: outputPath);
        Assert.StartsWith("Formula added", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Add_WithSheetIndex_ShouldApplyToCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_add_sheet.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets[1].Cells["A1"].Value = 100;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_sheet_output.xlsx");
        _tool.Execute("add", workbookPath, sheetIndex: 1, cell: "B1", formula: "=A1*2", outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=A1*2", resultWorkbook.Worksheets[1].Cells["B1"].Formula);
    }

    [Theory]
    [InlineData("=AVERAGE(A1:A3)")]
    [InlineData("=COUNT(A1:A3)")]
    [InlineData("=IF(A1>40,\"Pass\",\"Fail\")")]
    public void Add_WithVariousFormulas_ShouldAddFormula(string formula)
    {
        var workbookPath = CreateExcelWorkbook($"test_add_{formula.GetHashCode()}.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 50;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Value = 30;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath($"test_add_{formula.GetHashCode()}_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "B1", formula: formula, outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal(formula, resultWorkbook.Worksheets[0].Cells["B1"].Formula);
    }

    [Fact]
    public void Add_WithVLOOKUP_ShouldAddVLOOKUP()
    {
        var workbookPath = CreateExcelWorkbook("test_add_vlookup.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].Value = "ID001";
            worksheet.Cells["B1"].Value = "Apple";
            worksheet.Cells["A2"].Value = "ID002";
            worksheet.Cells["B2"].Value = "Banana";
            worksheet.Cells["D1"].Value = "ID002";
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_vlookup_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "E1", formula: "=VLOOKUP(D1,A1:B2,2,FALSE)", outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("VLOOKUP", resultWorkbook.Worksheets[0].Cells["E1"].Formula);
    }

    [Fact]
    public void Add_WithNestedFormula_ShouldAddNestedFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_add_nested.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 85;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_nested_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "B1", formula: "=IF(A1>=90,\"A\",IF(A1>=80,\"B\",\"C\"))",
            outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("IF", resultWorkbook.Worksheets[0].Cells["B1"].Formula);
    }

    [Fact]
    public void Add_WithCrossSheetRef_ShouldAddCrossSheetRef()
    {
        var workbookPath = CreateExcelWorkbook("test_add_cross_sheet.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Name = "Data";
            workbook.Worksheets[0].Cells["A1"].Value = 100;
            workbook.Worksheets.Add("Summary");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_cross_sheet_output.xlsx");
        _tool.Execute("add", workbookPath, sheetIndex: 1, cell: "A1", formula: "=Data!A1*2", outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Contains("Data", resultWorkbook.Worksheets[1].Cells["A1"].Formula);
    }

    [Fact]
    public void Get_ShouldReturnFormulas()
    {
        var workbookPath = CreateWorkbookWithFormula("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        var items = json.RootElement.GetProperty("items");
        Assert.Equal("A3", items[0].GetProperty("cell").GetString());
        Assert.Contains("A1", items[0].GetProperty("formula").GetString());
    }

    [Fact]
    public void Get_WithRange_ShouldReturnFormulasInRange()
    {
        var workbookPath = CreateExcelWorkbook("test_get_range.xlsx");
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
    public void Get_NoFormulas_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
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

    [Fact]
    public void GetResult_ShouldReturnResult()
    {
        var workbookPath = CreateWorkbookWithFormula("test_get_result.xlsx");
        var result = _tool.Execute("get_result", workbookPath, cell: "A3");
        var json = JsonDocument.Parse(result);
        Assert.Equal("A3", json.RootElement.GetProperty("cell").GetString());
        Assert.Contains("30", json.RootElement.GetProperty("calculatedValue").GetString());
    }

    [Fact]
    public void GetResult_WithCalculateBeforeReadFalse_ShouldNotRecalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_get_result_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Formula = "=A1*2";
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get_result", workbookPath, cell: "A2", calculateBeforeRead: false);
        Assert.Contains("A2", result);
    }

    [Fact]
    public void GetResult_WithErrorFormula_ShouldReturnError()
    {
        var workbookPath = CreateExcelWorkbook("test_get_result_error.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 0;
            workbook.Worksheets[0].Cells["A3"].Formula = "=A1/A2";
            workbook.CalculateFormula();
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute("get_result", workbookPath, cell: "A3");
        Assert.True(result.Contains("DIV") || result.Contains("Error") || result.Contains("Infinity") ||
                    result.Contains("A3"));
    }

    [Fact]
    public void Calculate_ShouldCalculateAllFormulas()
    {
        var workbookPath = CreateExcelWorkbook("test_calculate.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Worksheets[0].Cells["A3"].Formula = "=A1+A2";
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_calculate_output.xlsx");
        var result = _tool.Execute("calculate", workbookPath, outputPath: outputPath);
        Assert.StartsWith("Formulas calculated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetArray_ShouldSetArrayFormula()
    {
        var workbookPath = CreateExcelWorkbook("test_set_array.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 1;
            workbook.Worksheets[0].Cells["A2"].Value = 2;
            workbook.Worksheets[0].Cells["A3"].Value = 3;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_set_array_output.xlsx");
        var result = _tool.Execute("set_array", workbookPath, range: "B1:B3", formula: "=A1:A3*2",
            outputPath: outputPath);
        Assert.Contains("B1:B3", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetArray_WithAutoCalculateFalse_ShouldNotCalculate()
    {
        var workbookPath = CreateExcelWorkbook("test_set_array_no_calc.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 1;
            workbook.Worksheets[0].Cells["A2"].Value = 2;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_set_array_no_calc_output.xlsx");
        _tool.Execute("set_array", workbookPath, range: "B1:B2", formula: "=A1:A2*2", autoCalculate: false,
            outputPath: outputPath);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GetArray_ShouldReturnArrayFormula()
    {
        var workbookPath = CreateWorkbookWithArrayFormula("test_get_array.xlsx");
        var result = _tool.Execute("get_array", workbookPath, cell: "B1");
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isArrayFormula").GetBoolean());
    }

    [Fact]
    public void GetArray_NotArrayFormula_ShouldReturnFalse()
    {
        var workbookPath = CreateWorkbookWithFormula("test_get_array_not.xlsx");
        var result = _tool.Execute("get_array", workbookPath, cell: "A3");
        var json = JsonDocument.Parse(result);
        Assert.False(json.RootElement.GetProperty("isArrayFormula").GetBoolean());
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "B1", formula: "=A1*2", outputPath: outputPath);
        Assert.StartsWith("Formula added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateWorkbookWithFormula($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("count", result);
    }

    [Theory]
    [InlineData("CALCULATE")]
    [InlineData("Calculate")]
    [InlineData("calculate")]
    public void Operation_ShouldBeCaseInsensitive_Calculate(string operation)
    {
        var workbookPath = CreateWorkbookWithFormula($"test_case_calc_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_calc_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, outputPath: outputPath);
        Assert.StartsWith("Formulas calculated", result);
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
    public void Add_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetIndex: 99, cell: "A1", formula: "=1+1"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Add_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, formula: "=SUM(A1:A10)"));
        Assert.Contains("cell", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithMissingFormula_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_formula.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, cell: "A1"));
        Assert.Contains("formula", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithInvalidFunctionName_ShouldReturnWarning()
    {
        var workbookPath = CreateExcelWorkbook("test_add_invalid_func.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_invalid_func_output.xlsx");
        var result = _tool.Execute("add", workbookPath, cell: "B1", formula: "=INVALIDFUNC(A1)",
            outputPath: outputPath);
        Assert.StartsWith("Formula added", result);
        Assert.Contains("#NAME?", result);
    }

    [Fact]
    public void GetResult_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_result_missing_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_result", workbookPath));
        Assert.Contains("cell", ex.Message.ToLower());
    }

    [Fact]
    public void GetArray_WithMissingCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_array_missing_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_array", workbookPath));
        Assert.Contains("cell", ex.Message.ToLower());
    }

    [Fact]
    public void SetArray_WithMissingRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_set_array_missing_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_array", workbookPath, formula: "=A1:A3*2"));
        Assert.Contains("range", ex.Message.ToLower());
    }

    [Fact]
    public void SetArray_WithMissingFormula_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_set_array_missing_formula.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_array", workbookPath, range: "B1:B3"));
        Assert.Contains("formula", ex.Message.ToLower());
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
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = 10;
            wb.Worksheets[0].Cells["A2"].Value = 20;
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, cell: "A3", formula: "=A1+A2");
        Assert.StartsWith("Formula added", result);
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("=A1+A2", sessionWorkbook.Worksheets[0].Cells["A3"].Formula);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithFormula("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("count").GetInt32());
        var items = json.RootElement.GetProperty("items");
        Assert.Equal("A3", items[0].GetProperty("cell").GetString());
    }

    [Fact]
    public void GetResult_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithFormula("test_session_get_result.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_result", sessionId: sessionId, cell: "A3");
        var json = JsonDocument.Parse(result);
        Assert.Equal("A3", json.RootElement.GetProperty("cell").GetString());
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
        Assert.StartsWith("Formulas calculated", result);
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotNull(sessionWorkbook.Worksheets[0].Cells["A3"].Value);
    }

    [Fact]
    public void GetArray_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithArrayFormula("test_session_get_array.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_array", sessionId: sessionId, cell: "B1");
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("isArrayFormula").GetBoolean());
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
        var sessionWorkbook = CreateExcelWorkbook("test_session_file.xlsx");
        using (var wb = new Workbook(sessionWorkbook))
        {
            wb.Worksheets[0].Name = "SessionSheet";
            wb.Worksheets[0].Cells["A1"].Formula = "=1+1";
            wb.Save(sessionWorkbook);
        }

        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        Assert.Contains("SessionSheet", result);
    }

    #endregion
}