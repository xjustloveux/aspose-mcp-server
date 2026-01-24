using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Formula;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelFormulaTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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
        // CS0618 - Using obsolete SetArrayFormula for backward compatibility testing
#pragma warning disable CS0618
        workbook.Worksheets[0].Cells["B1"].SetArrayFormula("=A1:A2*2", 2, 1);
#pragma warning restore CS0618
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddFormulaAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_add.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Worksheets[0].Cells["A2"].Value = 20;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        _tool.Execute("add", workbookPath, cell: "A3", formula: "=A1+A2", outputPath: outputPath);
        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("=A1+A2", resultWorkbook.Worksheets[0].Cells["A3"].Formula);
    }

    [Fact]
    public void Get_ShouldReturnFormulasFromFile()
    {
        var workbookPath = CreateWorkbookWithFormula("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetFormulasResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Equal("A3", data.Items[0].Cell);
    }

    [Fact]
    public void GetResult_ShouldReturnFormulaResultFromFile()
    {
        var workbookPath = CreateWorkbookWithFormula("test_get_result.xlsx");
        var result = _tool.Execute("get_result", workbookPath, cell: "A3");
        var data = GetResultData<GetFormulaResultResult>(result);
        Assert.Equal("A3", data.Cell);
        Assert.Contains("30", data.CalculatedValue);
    }

    [Fact]
    public void Calculate_ShouldCalculateFormulasAndPersistToFile()
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Formulas calculated", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetArray_ShouldSetArrayFormulaAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_set_array.xlsx");
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 1;
            workbook.Worksheets[0].Cells["A2"].Value = 2;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_set_array_output.xlsx");
        var result = _tool.Execute("set_array", workbookPath, range: "B1:B2", formula: "=A1:A2*2",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("B1:B2", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GetArray_ShouldReturnArrayFormulaInfoFromFile()
    {
        var workbookPath = CreateWorkbookWithArrayFormula("test_get_array.xlsx");
        var result = _tool.Execute("get_array", workbookPath, cell: "B1");
        var data = GetResultData<GetArrayFormulaResult>(result);
        Assert.True(data.IsArrayFormula);
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
        using (var workbook = new Workbook(workbookPath))
        {
            workbook.Worksheets[0].Cells["A1"].Value = 10;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, cell: "B1", formula: "=A1*2", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Formula added", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session Management

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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Formula added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("=A1+A2", sessionWorkbook.Worksheets[0].Cells["A3"].Formula);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithFormula("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetFormulasResult>(result);
        Assert.Equal(1, data.Count);
        var output = GetResultOutput<GetFormulasResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void GetResult_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithFormula("test_session_get_result.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_result", sessionId: sessionId, cell: "A3");
        var data = GetResultData<GetFormulaResultResult>(result);
        Assert.Equal("A3", data.Cell);
        var output = GetResultOutput<GetFormulaResultResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Formulas calculated", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var sessionWorkbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotNull(sessionWorkbook.Worksheets[0].Cells["A3"].Value);
    }

    [Fact]
    public void GetArray_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithArrayFormula("test_session_get_array.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_array", sessionId: sessionId, cell: "B1");
        var data = GetResultData<GetArrayFormulaResult>(result);
        Assert.True(data.IsArrayFormula);
        var output = GetResultOutput<GetArrayFormulaResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<GetFormulasResult>(result);
        Assert.Equal("SessionSheet", data.WorksheetName);
    }

    #endregion
}
