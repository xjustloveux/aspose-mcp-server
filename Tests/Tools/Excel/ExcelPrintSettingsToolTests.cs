using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelPrintSettingsToolTests : ExcelTestBase
{
    private readonly ExcelPrintSettingsTool _tool;

    public ExcelPrintSettingsToolTests()
    {
        _tool = new ExcelPrintSettingsTool(SessionManager);
    }

    #region General

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
    public void SetPrintArea_MultipleRanges_ShouldSetMultiplePrintAreas()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_area_multi.xlsx", 10, 10);
        var outputPath = CreateTestFilePath("test_set_print_area_multi_output.xlsx");
        var result = _tool.Execute("set_print_area", workbookPath, range: "A1:D10,F1:H10", outputPath: outputPath);
        Assert.Contains("Print area", result);
        using var workbook = new Workbook(outputPath);
        Assert.Contains("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
        Assert.Contains("F1:H10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    [Fact]
    public void SetPrintArea_ClearPrintArea_ShouldClearPrintArea()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_clear_print_area.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].PageSetup.PrintArea = "A1:D10";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_clear_print_area_output.xlsx");
        var result = _tool.Execute("set_print_area", workbookPath, clearPrintArea: true, outputPath: outputPath);
        Assert.Contains("cleared", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintArea));
    }

    [Fact]
    public void SetPrintArea_WithSheetIndex_ShouldSetPrintAreaOnCorrectSheet()
    {
        var workbookPath = CreateExcelWorkbook("test_print_area_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells["A1"].PutValue("Test");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_print_area_sheet_output.xlsx");
        var result = _tool.Execute("set_print_area", workbookPath, sheetIndex: 1, range: "A1:B5",
            outputPath: outputPath);
        Assert.Contains("Print area", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("A1:B5", workbook.Worksheets[1].PageSetup.PrintArea);
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
    public void SetPrintTitles_OnlyRows_ShouldSetOnlyRows()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_print_titles_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_print_titles_rows_output.xlsx");
        var result = _tool.Execute("set_print_titles", workbookPath, rows: "$1:$2", outputPath: outputPath);
        Assert.Contains("titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
    }

    [Fact]
    public void SetPrintTitles_OnlyColumns_ShouldSetOnlyColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_print_titles_cols.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_print_titles_cols_output.xlsx");
        var result = _tool.Execute("set_print_titles", workbookPath, columns: "$A:$B", outputPath: outputPath);
        Assert.Contains("titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$A:$B", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public void SetPrintTitles_ClearTitles_ShouldClearPrintTitles()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_clear_print_titles.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].PageSetup.PrintTitleRows = "$1:$1";
            wb.Worksheets[0].PageSetup.PrintTitleColumns = "$A:$A";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_clear_print_titles_output.xlsx");
        var result = _tool.Execute("set_print_titles", workbookPath, clearTitles: true, outputPath: outputPath);
        Assert.Contains("titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintTitleRows));
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintTitleColumns));
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
    public void SetPageSetup_WithMargins_ShouldSetMargins()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_margins.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_margins_output.xlsx");
        var result = _tool.Execute("set_page_setup", workbookPath,
            leftMargin: 0.5, rightMargin: 0.5, topMargin: 0.75, bottomMargin: 0.75, outputPath: outputPath);
        Assert.Contains("Page setup updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(0.5, workbook.Worksheets[0].PageSetup.LeftMargin, 5);
        Assert.Equal(0.5, workbook.Worksheets[0].PageSetup.RightMargin, 5);
        Assert.Equal(0.75, workbook.Worksheets[0].PageSetup.TopMargin, 5);
        Assert.Equal(0.75, workbook.Worksheets[0].PageSetup.BottomMargin, 5);
    }

    [Fact]
    public void SetPageSetup_WithHeaderFooter_ShouldSetHeaderFooter()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_header_footer.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_header_footer_output.xlsx");
        var result = _tool.Execute("set_page_setup", workbookPath, header: "Test Header", footer: "Page &P of &N",
            outputPath: outputPath);
        Assert.Contains("Page setup updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Contains("Test Header", workbook.Worksheets[0].PageSetup.GetHeader(1));
        Assert.Contains("Page", workbook.Worksheets[0].PageSetup.GetFooter(1));
    }

    [Fact]
    public void SetPageSetup_WithFitToPage_ShouldSetFitToPage()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_fit.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_fit_output.xlsx");
        var result = _tool.Execute("set_page_setup", workbookPath, fitToPage: true, fitToPagesWide: 1,
            fitToPagesTall: 0, outputPath: outputPath);
        Assert.Contains("Page setup updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(1, workbook.Worksheets[0].PageSetup.FitToPagesWide);
        Assert.Equal(0, workbook.Worksheets[0].PageSetup.FitToPagesTall);
    }

    [Fact]
    public void SetPageSetup_NoChanges_ShouldReturnNoChangesMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_no_changes.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_no_changes_output.xlsx");
        var result = _tool.Execute("set_page_setup", workbookPath, outputPath: outputPath);
        Assert.Contains("Page setup updated", result);
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
        Assert.Equal(PageOrientationType.Portrait, workbook.Worksheets[0].PageSetup.Orientation);
        Assert.Equal(PaperSizeType.PaperA4, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Fact]
    public void SetAll_WithPrintTitles_ShouldSetPrintTitles()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_titles.xlsx", 20, 10);
        var outputPath = CreateTestFilePath("test_set_all_titles_output.xlsx");
        var result = _tool.Execute("set_all", workbookPath, range: "A1:J20", rows: "$1:$2", columns: "$A:$A",
            outputPath: outputPath);
        Assert.Contains("Print settings updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public void SetAll_NoChanges_ShouldReturnNoChangesMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_set_all_no_changes.xlsx");
        var outputPath = CreateTestFilePath("test_set_all_no_changes_output.xlsx");
        var result = _tool.Execute("set_all", workbookPath, outputPath: outputPath);
        Assert.Contains("Print settings updated", result);
    }

    [Fact]
    public void SetPageSetup_DefaultOutputPath_ShouldUseInputPath()
    {
        var workbookPath = CreateExcelWorkbook("test_default_output.xlsx");
        var result = _tool.Execute("set_page_setup", workbookPath, orientation: "Landscape");
        Assert.Contains("Page setup updated", result);
        using var workbook = new Workbook(workbookPath);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
    }

    [Theory]
    [InlineData("SET_PRINT_AREA")]
    [InlineData("Set_Print_Area")]
    [InlineData("set_print_area")]
    public void Operation_ShouldBeCaseInsensitive_SetPrintArea(string operation)
    {
        var workbookPath = CreateExcelWorkbookWithData($"test_case_{operation.Replace("_", "")}.xlsx", 5, 5);
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, range: "A1:C3", outputPath: outputPath);
        Assert.Contains("Print area", result);
    }

    [Theory]
    [InlineData("SET_PAGE_SETUP")]
    [InlineData("Set_Page_Setup")]
    [InlineData("set_page_setup")]
    public void Operation_ShouldBeCaseInsensitive_SetPageSetup(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_{operation.Replace("_", "")}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, orientation: "Landscape", outputPath: outputPath);
        Assert.Contains("Page setup updated", result);
    }

    [Theory]
    [InlineData("letter", PaperSizeType.PaperLetter)]
    [InlineData("LETTER", PaperSizeType.PaperLetter)]
    [InlineData("a4", PaperSizeType.PaperA4)]
    [InlineData("A4", PaperSizeType.PaperA4)]
    public void SetPageSetup_CaseInsensitivePaperSize_ShouldWork(string paperSize, PaperSizeType expected)
    {
        var workbookPath = CreateExcelWorkbook($"test_paper_{paperSize}.xlsx");
        var outputPath = CreateTestFilePath($"test_paper_{paperSize}_output.xlsx");
        _tool.Execute("set_page_setup", workbookPath, paperSize: paperSize, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(expected, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Theory]
    [InlineData("landscape", PageOrientationType.Landscape)]
    [InlineData("LANDSCAPE", PageOrientationType.Landscape)]
    [InlineData("portrait", PageOrientationType.Portrait)]
    [InlineData("PORTRAIT", PageOrientationType.Portrait)]
    public void SetPageSetup_CaseInsensitiveOrientation_ShouldWork(string orientation, PageOrientationType expected)
    {
        var workbookPath = CreateExcelWorkbook($"test_orient_{orientation}.xlsx");
        var outputPath = CreateTestFilePath($"test_orient_{orientation}_output.xlsx");
        _tool.Execute("set_page_setup", workbookPath, orientation: orientation, outputPath: outputPath);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(expected, workbook.Worksheets[0].PageSetup.Orientation);
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
    public void SetPrintArea_WithoutRangeOrClear_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_print_area_no_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("set_print_area", workbookPath));
        Assert.Contains("Either range or clearPrintArea must be provided", ex.Message);
    }

    [Fact]
    public void SetPrintArea_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_print_area_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_print_area", workbookPath, sheetIndex: 99, range: "A1:B5"));
    }

    [Fact]
    public void SetPrintTitles_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_print_titles_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_print_titles", workbookPath, sheetIndex: 99, rows: "$1:$1"));
    }

    [Fact]
    public void SetPageSetup_InvalidPaperSize_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_invalid_paper.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_page_setup", workbookPath, paperSize: "InvalidSize"));
        Assert.Contains("Invalid paper size", ex.Message);
        Assert.Contains("Supported values", ex.Message);
    }

    [Fact]
    public void SetPageSetup_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_page_setup", workbookPath, sheetIndex: 99, orientation: "Landscape"));
    }

    [Fact]
    public void SetAll_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_set_all_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set_all", workbookPath, sheetIndex: 99, range: "A1:D10"));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("set_print_area", range: "A1:D10"));
    }

    #endregion

    #region Session

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
    public void SetPrintTitles_WithSessionId_ShouldSetInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_print_titles.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("set_print_titles", sessionId: sessionId, rows: "$1:$2", columns: "$A:$A");
        Assert.Contains("titles updated", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
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
        Assert.Equal(PaperSizeType.PaperA4, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Fact]
    public void SetAll_WithSessionId_ShouldSetInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_set_all.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("set_all", sessionId: sessionId, range: "A1:D10", orientation: "Landscape");
        Assert.Contains("Print settings updated", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
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