using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelRangeTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelRangeToolTests : ExcelTestBase
{
    private readonly ExcelRangeTool _tool;

    public ExcelRangeToolTests()
    {
        _tool = new ExcelRangeTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Write_ShouldWriteDataAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_write.xlsx");
        var outputPath = CreateTestFilePath("test_write_output.xlsx");
        var data = "[[\"A\", \"B\"], [\"C\", \"D\"]]";
        _tool.Execute("write", workbookPath, startCell: "A1", data: data, outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal("A", worksheet.Cells["A1"].Value);
        Assert.Equal("D", worksheet.Cells["B2"].Value);
    }

    [Fact]
    public void Get_ShouldReturnRangeDataFromFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get.xlsx", 3);
        var result = _tool.Execute("get", workbookPath, range: "A1:B2");
        Assert.NotNull(result);
        Assert.Contains("R1C1", result);
    }

    [Fact]
    public void Edit_ShouldEditRangeAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit.xlsx", 3);
        var outputPath = CreateTestFilePath("test_edit_output.xlsx");
        var data = "[[\"X\", \"Y\"], [\"Z\", \"W\"]]";
        _tool.Execute("edit", workbookPath, range: "A1:B2", data: data, outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        Assert.Equal("X", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Clear_ShouldClearRangeAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_clear.xlsx", 3);
        var outputPath = CreateTestFilePath("test_clear_output.xlsx");
        _tool.Execute("clear", workbookPath, range: "A1:B2", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var a1Value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(a1Value == null || a1Value.ToString() == "");
    }

    [Fact]
    public void Copy_ShouldCopyRangeAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_copy.xlsx", 3);
        var outputPath = CreateTestFilePath("test_copy_output.xlsx");
        _tool.Execute("copy", workbookPath, sourceRange: "A1:B2", destCell: "C1", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.Equal(worksheet.Cells["A1"].Value?.ToString() ?? "", worksheet.Cells["C1"].Value?.ToString() ?? "");
    }

    [Fact]
    public void Move_ShouldMoveRangeAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_move.xlsx", 3);
        var originalWorkbook = new Workbook(workbookPath);
        var sourceA1Value = originalWorkbook.Worksheets[0].Cells["A1"].Value;

        var outputPath = CreateTestFilePath("test_move_output.xlsx");
        _tool.Execute("move", workbookPath, sourceRange: "A1:B2", destCell: "C1", outputPath: outputPath);
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];

        Assert.Equal(sourceA1Value, worksheet.Cells["C1"].Value);
        var a1 = worksheet.Cells["A1"].Value;
        Assert.True(a1 == null || a1.ToString() == "", "A1 should be cleared after move");
    }

    [Fact]
    public void CopyFormat_ShouldCopyFormatAndPersistToFile()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_format.xlsx");
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        var sourceCell = worksheet.Cells["A1"];
        sourceCell.Value = "Test";
        var style = sourceCell.GetStyle();
        style.Font.IsBold = true;
        sourceCell.SetStyle(style);
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_format_output.xlsx");
        _tool.Execute("copy_format", workbookPath, sourceRange: "A1", destCell: "B1", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var destStyle = resultWorkbook.Worksheets[0].Cells["B1"].GetStyle();
        Assert.True(destStyle.Font.IsBold);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("WRITE")]
    [InlineData("Write")]
    [InlineData("write")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var data = "[[\"Test\"]]";
        var result = _tool.Execute(operation, workbookPath, startCell: "A1", data: data, outputPath: outputPath);
        Assert.Contains("A1", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get", range: "A1"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Write_WithSessionId_ShouldWriteInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_write.xlsx");
        var sessionId = OpenSession(workbookPath);
        var data = "[[\"SessionA\", \"SessionB\"], [\"SessionC\", \"SessionD\"]]";
        var result = _tool.Execute("write", sessionId: sessionId, startCell: "A1", data: data);
        Assert.StartsWith("Data written", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("SessionA", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId, range: "A1:B2");
        Assert.NotNull(result);
        Assert.Contains("R1C1", result);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_edit.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var data = "[[\"X\", \"Y\"]]";
        var result = _tool.Execute("edit", sessionId: sessionId, range: "A1:B1", data: data);
        Assert.StartsWith("Range A1:B1 edited", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("X", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_clear.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("clear", sessionId: sessionId, range: "A1:B2");
        Assert.StartsWith("Range A1:B2 cleared", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var a1Value = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(a1Value == null || a1Value.ToString() == "");
    }

    [Fact]
    public void Copy_WithSessionId_ShouldCopyInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_copy.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("copy", sessionId: sessionId, sourceRange: "A1:B2", destCell: "D1");
        Assert.StartsWith("Range A1:B2 copied", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var worksheet = workbook.Worksheets[0];
        var sourceA1 = worksheet.Cells["A1"].Value?.ToString() ?? "";
        var destD1 = worksheet.Cells["D1"].Value?.ToString() ?? "";
        Assert.Equal(sourceA1, destD1);
    }

    [Fact]
    public void Move_WithSessionId_ShouldMoveInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_move.xlsx", 3);
        var sessionId = OpenSession(workbookPath);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var originalValue = workbook.Worksheets[0].Cells["A1"].Value;
        var result = _tool.Execute("move", sessionId: sessionId, sourceRange: "A1:B2", destCell: "D1");
        Assert.StartsWith("Range A1:B2 moved", result);
        Assert.Equal(originalValue, workbook.Worksheets[0].Cells["D1"].Value);
        var a1 = workbook.Worksheets[0].Cells["A1"].Value;
        Assert.True(a1 == null || a1.ToString() == "");
    }

    [Fact]
    public void CopyFormat_WithSessionId_ShouldCopyInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_copy_format.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            var style = wb.CreateStyle();
            style.Font.IsBold = true;
            wb.Worksheets[0].Cells["A1"].SetStyle(style);
            wb.Worksheets[0].Cells["A1"].Value = "Test";
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("copy_format", sessionId: sessionId, range: "A1", destCell: "B1");
        Assert.StartsWith("Format copied", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var destStyle = workbook.Worksheets[0].Cells["B1"].GetStyle();
        Assert.True(destStyle.Font.IsBold);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session", range: "A1"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbookWithData("test_path_file.xlsx", 2);
        var workbookPath2 = CreateExcelWorkbookWithData("test_session_file.xlsx");
        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId, range: "A1:C5");
        Assert.Contains("R5C3", result);
    }

    #endregion
}
