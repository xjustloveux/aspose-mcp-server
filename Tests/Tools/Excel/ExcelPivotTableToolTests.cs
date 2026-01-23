using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.PivotTable;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelPivotTableTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelPivotTableToolTests : ExcelTestBase
{
    private readonly ExcelPivotTableTool _tool;

    public ExcelPivotTableToolTests()
    {
        _tool = new ExcelPivotTableTool(SessionManager);
    }

    private string CreateWorkbookWithPivotTable(string fileName)
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[0, 0].Value = "Category";
        worksheet.Cells[0, 1].Value = "Sales";
        worksheet.Cells[0, 2].Value = "Region";
        worksheet.Cells[0, 3].Value = "Quantity";
        for (var row = 1; row <= 10; row++)
        {
            worksheet.Cells[row, 0].Value = $"Cat{row % 3}";
            worksheet.Cells[row, 1].Value = row * 100;
            worksheet.Cells[row, 2].Value = $"Region{row % 2}";
            worksheet.Cells[row, 3].Value = row * 10;
        }

        worksheet.PivotTables.Add("A1:D11", "F1", "PivotTable1");
        workbook.Save(workbookPath);
        return workbookPath;
    }

    private string CreateWorkbookWithDataForPivot(string fileName)
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[0, 0].Value = "Category";
        worksheet.Cells[0, 1].Value = "Sales";
        worksheet.Cells[0, 2].Value = "Region";
        worksheet.Cells[0, 3].Value = "Quantity";
        for (var row = 1; row <= 10; row++)
        {
            worksheet.Cells[row, 0].Value = $"Cat{row % 3}";
            worksheet.Cells[row, 1].Value = row * 100;
            worksheet.Cells[row, 2].Value = $"Region{row % 2}";
            worksheet.Cells[row, 3].Value = row * 10;
        }

        workbook.Save(workbookPath);
        return workbookPath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddPivotTableAndPersistToFile()
    {
        var workbookPath = CreateWorkbookWithDataForPivot("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, sourceRange: "A1:D11", destCell: "F1", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].PivotTables.Count > 0);
    }

    [Fact]
    public void Get_ShouldReturnPivotTableInfoFromFile()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetPivotTablesResult>(result);
        Assert.Equal(1, data.Count);
    }

    [Fact]
    public void Delete_ShouldDeletePivotTableAndPersistToFile()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, pivotTableIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].PivotTables);
    }

    [Fact]
    public void Edit_ShouldEditPivotTableAndPersistToFile()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit.xlsx");
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            name: "EditedPivot", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("edited", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("EditedPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void AddField_ShouldAddFieldToPivotTable()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_add_field.xlsx");
        var outputPath = CreateTestFilePath("test_add_field_output.xlsx");
        var result = _tool.Execute("add_field", workbookPath, pivotTableIndex: 0,
            fieldName: "Region", area: "Row", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message);
    }

    [Fact]
    public void Refresh_ShouldRefreshPivotTable()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_refresh.xlsx");
        var outputPath = CreateTestFilePath("test_refresh_output.xlsx");
        var result = _tool.Execute("refresh", workbookPath, pivotTableIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Refreshed", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateWorkbookWithDataForPivot($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, sourceRange: "A1:D11", destCell: "F1",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message);
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
        var workbookPath = CreateWorkbookWithDataForPivot("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            sourceRange: "A1:D11", destCell: "F1", name: "SessionPivot");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].PivotTables.Count > 0);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetPivotTablesResult>(result);
        Assert.Equal(1, data.Count);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete", sessionId: sessionId, pivotTableIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].PivotTables);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_edit.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("edit", sessionId: sessionId, pivotTableIndex: 0, name: "UpdatedPivot");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("edited", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("UpdatedPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void AddField_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_addfield.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add_field", sessionId: sessionId, pivotTableIndex: 0,
            fieldName: "Region", area: "Row");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Refresh_WithSessionId_ShouldRefreshInMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_refresh.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("refresh", sessionId: sessionId, pivotTableIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Refreshed", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
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
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateWorkbookWithPivotTable("test_session_file.xlsx");
        using (var wb = new Workbook(workbookPath2))
        {
            wb.Worksheets[0].PivotTables[0].Name = "SessionPivotTable";
            wb.Save(workbookPath2);
        }

        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        var data = GetResultData<GetPivotTablesResult>(result);
        Assert.Contains(data.Items, pt => pt.Name == "SessionPivotTable");
    }

    #endregion
}
