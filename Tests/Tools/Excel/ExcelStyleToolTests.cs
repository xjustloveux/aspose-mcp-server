using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelStyleTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelStyleToolTests : ExcelTestBase
{
    private readonly ExcelStyleTool _tool;

    public ExcelStyleToolTests()
    {
        _tool = new ExcelStyleTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void FormatCells_WithFontOptions_ShouldApplyFontFormatting()
    {
        var workbookPath = CreateExcelWorkbook("test_format_font.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_font_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", fontName: "Arial", fontSize: 14,
            bold: true, italic: true, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal("Arial", style.Font.Name);
        Assert.Equal(14, style.Font.Size);
        Assert.True(style.Font.IsBold);
        Assert.True(style.Font.IsItalic);
    }

    [Fact]
    public void GetFormat_ShouldReturnFormatInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format.xlsx");
        var workbook = new Workbook(workbookPath);
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Font.Name = "Arial";
        style.Font.Size = 14;
        style.Font.IsBold = true;
        cell.SetStyle(style);
        workbook.Save(workbookPath);

        var result = _tool.Execute("get_format", workbookPath, range: "A1");
        Assert.Contains("A1", result);
        Assert.Contains("fontName", result);
    }

    [Fact]
    public void CopySheetFormat_ShouldCopyFormat()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_format.xlsx");
        var workbook = new Workbook(workbookPath);
        var sourceSheet = workbook.Worksheets[0];
        sourceSheet.Cells["A1"].Value = "Test";
        sourceSheet.Cells.SetColumnWidth(0, 20);
        workbook.Worksheets.Add("TargetSheet");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_format_output.xlsx");
        var result = _tool.Execute("copy_sheet_format", workbookPath, sourceSheetIndex: 0,
            targetSheetIndex: 1, copyColumnWidths: true, copyRowHeights: true, outputPath: outputPath);
        Assert.Contains("Sheet format copied", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("FORMAT")]
    [InlineData("Format")]
    [InlineData("format")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation}.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        _tool.Execute(operation, workbookPath, range: "A1", bold: true, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(resultWorkbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
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
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_format", range: "A1"));
    }

    #endregion

    #region Session Management

    [Fact]
    public void GetFormat_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_get_format.xlsx");
        var workbook = new Workbook(workbookPath);
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Font.IsBold = true;
        cell.SetStyle(style);
        workbook.Save(workbookPath);

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get_format", sessionId: sessionId, range: "A1");
        Assert.Contains("A1", result);
    }

    [Fact]
    public void Format_WithSessionId_ShouldFormatInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_format.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells["A1"].Value = "Test";
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        _tool.Execute("format", sessionId: sessionId, range: "A1", bold: true, fontColor: "#FF0000");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.True(style.Font.IsBold);
    }

    [SkippableFact]
    public void CopySheetFormat_WithSessionId_ShouldCopyInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Evaluation mode adds extra watermark worksheets");

        var workbookPath = CreateExcelWorkbook("test_session_copy.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells.SetColumnWidth(0, 25);
            wb.Worksheets.Add("TargetSheet");
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("copy_sheet_format", sessionId: sessionId, sourceSheetIndex: 0,
            targetSheetIndex: 1, copyColumnWidths: true);
        Assert.Contains("Sheet format copied", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(2, workbook.Worksheets.Count);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_format", sessionId: "invalid_session", range: "A1"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateExcelWorkbook("test_session_file.xlsx");
        using (var wb = new Workbook(workbookPath2))
        {
            var cell = wb.Worksheets[0].Cells["A1"];
            cell.Value = "SessionValue";
            var style = cell.GetStyle();
            style.Font.IsBold = true;
            cell.SetStyle(style);
            wb.Save(workbookPath2);
        }

        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get_format", workbookPath1, sessionId, range: "A1");
        Assert.Contains("A1", result);

        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
    }

    #endregion
}
