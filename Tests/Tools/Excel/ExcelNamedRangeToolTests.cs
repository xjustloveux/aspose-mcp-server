using Aspose.Cells;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.Excel.NamedRange;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelNamedRangeTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelNamedRangeToolTests : ExcelTestBase
{
    private readonly ExcelNamedRangeTool _tool;

    public ExcelNamedRangeToolTests()
    {
        _tool = new ExcelNamedRangeTool(SessionManager);
    }

    private string CreateWorkbookWithNamedRange(string fileName, string rangeName, string rangeAddress)
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        for (var row = 0; row < 5; row++)
        for (var col = 0; col < 5; col++)
            worksheet.Cells[row, col].Value = $"R{row}C{col}";
        var parts = rangeAddress.Split(':');
        var range = parts.Length > 1
            ? worksheet.Cells.CreateRange(parts[0], parts[1])
            : worksheet.Cells.CreateRange(parts[0], parts[0]);
        range.Name = rangeName;
        workbook.Save(workbookPath);
        return workbookPath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddNamedRange()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add.xlsx", 5, 5);
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, name: "TestRange", range: "A1:C5", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Named range 'TestRange' added", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.NotNull(workbook.Worksheets.Names["TestRange"]);
    }

    [Fact]
    public void Delete_ShouldDeleteNamedRange()
    {
        var workbookPath = CreateWorkbookWithNamedRange("test_delete.xlsx", "RangeToDelete", "A1:B2");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, name: "RangeToDelete", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Named range 'RangeToDelete' deleted", data.Message);
        using var workbook = new Workbook(outputPath);
        Assert.Null(workbook.Worksheets.Names["RangeToDelete"]);
    }

    [Fact]
    public void Get_ShouldReturnAllNamedRanges()
    {
        var workbookPath = CreateExcelWorkbook("test_get.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells.CreateRange("A1", "B2").Name = "Range1";
            wb.Worksheets[0].Cells.CreateRange("C1", "D2").Name = "Range2";
            wb.Save(workbookPath);
        }

        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetNamedRangesResult>(result);
        Assert.Equal(2, data.Count);
    }

    [Fact]
    public void Get_NoNamedRanges_ShouldReturnEmptyMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var data = GetResultData<GetNamedRangesResult>(result);
        Assert.Equal(0, data.Count);
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
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, name: $"Range_{operation}",
            range: "A1:B2", outputPath: outputPath);
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

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_add.xlsx", 5, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId, name: "InMemoryRange", range: "A1:C3");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Named range 'InMemoryRange' added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotNull(workbook.Worksheets.Names["InMemoryRange"]);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithNamedRange("test_session_delete.xlsx", "RangeToDelete", "A1:B2");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete", sessionId: sessionId, name: "RangeToDelete");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Named range 'RangeToDelete' deleted", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Null(workbook.Worksheets.Names["RangeToDelete"]);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithNamedRange("test_session_get.xlsx", "SessionRange", "A1:B2");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetNamedRangesResult>(result);
        Assert.Equal(1, data.Count);
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
        var workbookPath2 = CreateWorkbookWithNamedRange("test_session_file.xlsx", "SessionRange", "A1:B2");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        var data = GetResultData<GetNamedRangesResult>(result);
        Assert.Equal(1, data.Count);
        Assert.Contains(data.Items, i => i.Name == "SessionRange");
    }

    #endregion
}
