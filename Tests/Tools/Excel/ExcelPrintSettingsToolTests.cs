using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

/// <summary>
///     Integration tests for ExcelPrintSettingsTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class ExcelPrintSettingsToolTests : ExcelTestBase
{
    private readonly ExcelPrintSettingsTool _tool;

    public ExcelPrintSettingsToolTests()
    {
        _tool = new ExcelPrintSettingsTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void SetPrintArea_ShouldSetPrintArea()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_area.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_print_area_output.xlsx");
        var result = _tool.Execute("set_print_area", workbookPath, range: "A1:D10", outputPath: outputPath);
        Assert.Contains("Print area", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    [Fact]
    public void SetPrintTitles_ShouldSetBothRowsAndColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_titles.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_print_titles_output.xlsx");
        var result = _tool.Execute("set_print_titles", workbookPath, rows: "$1:$1", columns: "$A:$A",
            outputPath: outputPath);
        Assert.Contains("titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$1", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public void SetPageSetup_ShouldSetOrientationAndPaperSize()
    {
        var workbookPath = CreateExcelWorkbook("test_set_page_setup.xlsx");
        var outputPath = CreateTestFilePath("test_set_page_setup_output.xlsx");
        var result = _tool.Execute("set_page_setup", workbookPath, orientation: "Landscape", paperSize: "A4",
            outputPath: outputPath);
        Assert.Contains("Page setup updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
        Assert.Equal(PaperSizeType.PaperA4, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Fact]
    public void SetAll_ShouldSetAllPrintSettings()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_all.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_all_output.xlsx");
        var result = _tool.Execute("set_all", workbookPath, range: "A1:D10", orientation: "Portrait", paperSize: "A4",
            outputPath: outputPath);
        Assert.Contains("Print settings updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SET_PRINT_AREA")]
    [InlineData("Set_Print_Area")]
    [InlineData("set_print_area")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation.Replace("_", "")}.xlsx", 5, 5);
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:C3", outputPath: outputPath);
        Assert.Contains("Print area", result);
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
    public void SetPrintArea_WithSessionId_ShouldSetInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_print_area.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("set_print_area", sessionId: sessionId, range: "A1:D10");
        Assert.Contains("Print area", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    [Fact]
    public void SetPageSetup_WithSessionId_ShouldSetInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_page_setup.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("set_page_setup", sessionId: sessionId, orientation: "Landscape", paperSize: "A4");
        Assert.Contains("Page setup updated", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("set_print_area", sessionId: "invalid_session", range: "A1:D10"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateExcelWorkbook("test_session_file.xlsx");
        var sessionId = OpenSession(workbookPath2);
        _tool.Execute("set_print_area", workbookPath1, sessionId, range: "A1:D10");
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    #endregion
}
