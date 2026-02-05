using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.Table;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelTableTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelTableToolTests : ExcelTestBase
{
    private readonly ExcelTableTool _tool;

    public ExcelTableToolTests()
    {
        _tool = new ExcelTableTool(SessionManager);
    }

    private string CreateWorkbookWithData(string fileName)
    {
        var path = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["B1"].Value = "Value";
        sheet.Cells["C1"].Value = "Count";
        sheet.Cells["A2"].Value = "A";
        sheet.Cells["B2"].Value = 1;
        sheet.Cells["C2"].Value = 10;
        sheet.Cells["A3"].Value = "B";
        sheet.Cells["B3"].Value = 2;
        sheet.Cells["C3"].Value = 20;
        workbook.Save(path);
        return path;
    }

    private string CreateWorkbookWithTable(string fileName)
    {
        var path = CreateWorkbookWithData(fileName);
        using var workbook = new Workbook(path);
        workbook.Worksheets[0].ListObjects.Add("A1", "C3", true);
        workbook.Save(path);
        return path;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Create_ShouldCreateTable()
    {
        var workbookPath = CreateWorkbookWithData("test_create.xlsx");
        var outputPath = CreateTestFilePath("test_create_output.xlsx");
        var result = _tool.Execute("create", workbookPath, range: "A1:C3", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Table", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Single(workbook.Worksheets[0].ListObjects);
    }

    [Fact]
    public void Get_ShouldReturnTables()
    {
        var workbookPath = CreateWorkbookWithTable("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetTablesExcelResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Single(data.Items);
    }

    [Fact]
    public void Delete_ShouldDeleteTable()
    {
        var workbookPath = CreateWorkbookWithTable("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, tableIndex: 0, keepData: false, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("delete", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].ListObjects);
    }

    [Fact]
    public void SetStyle_ShouldSetStyle()
    {
        var workbookPath = CreateWorkbookWithTable("test_set_style.xlsx");
        var outputPath = CreateTestFilePath("test_set_style_output.xlsx");
        var result = _tool.Execute("set_style", workbookPath, tableIndex: 0, styleName: "TableStyleMedium9",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("style", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddTotalRow_ShouldAddTotalRow()
    {
        var workbookPath = CreateWorkbookWithTable("test_total_row.xlsx");
        var outputPath = CreateTestFilePath("test_total_row_output.xlsx");
        var result = _tool.Execute("add_total_row", workbookPath, tableIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("total", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].ListObjects[0].ShowTotals);
    }

    [Fact]
    public void ConvertToRange_ShouldConvert()
    {
        var workbookPath = CreateWorkbookWithTable("test_convert.xlsx");
        var outputPath = CreateTestFilePath("test_convert_output.xlsx");
        var result = _tool.Execute("convert_to_range", workbookPath, tableIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("convert", data.Message, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].ListObjects);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("CREATE")]
    [InlineData("Create")]
    [InlineData("create")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateWorkbookWithData($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:C3", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Table", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithData("test_unknown_op.xlsx");
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
    public void Create_WithSession_ShouldCreateInMemory()
    {
        var workbookPath = CreateWorkbookWithData("test_session_create.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("create", sessionId: sessionId, range: "A1:C3");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("Table", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Single(workbook.Worksheets[0].ListObjects);
    }

    [Fact]
    public void Get_WithSession_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithTable("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetTablesExcelResult>(result);
        Assert.Equal(1, data.Count);
        var output = GetResultOutput<GetTablesExcelResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Delete_WithSession_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithTable("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        _tool.Execute("delete", sessionId: sessionId, tableIndex: 0, keepData: false);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].ListObjects);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pathWorkbook = CreateWorkbookWithData("test_path_file.xlsx");
        var sessionWorkbook = CreateWorkbookWithTable("test_session_file.xlsx");
        var sessionId = OpenSession(sessionWorkbook);
        var result = _tool.Execute("get", pathWorkbook, sessionId);
        var data = GetResultData<GetTablesExcelResult>(result);
        Assert.Equal(1, data.Count);
    }

    #endregion
}
