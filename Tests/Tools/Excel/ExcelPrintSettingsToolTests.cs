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

    #region General Tests

    #region SetPrintArea Tests

    [Fact]
    public void SetPrintArea_ShouldSetPrintArea()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_area.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_print_area_output.xlsx");
        var result = _tool.Execute(
            "set_print_area",
            workbookPath,
            range: "A1:D10",
            outputPath: outputPath);
        Assert.Contains("Print area set to A1:D10", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    [Fact]
    public void SetPrintArea_MultipleRanges_ShouldSetMultiplePrintAreas()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_area_multi.xlsx", 10, 10);
        var outputPath = CreateTestFilePath("test_set_print_area_multi_output.xlsx");
        var result = _tool.Execute(
            "set_print_area",
            workbookPath,
            range: "A1:D10,F1:H10",
            outputPath: outputPath);
        Assert.Contains("Print area set to A1:D10,F1:H10", result);
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
        var result = _tool.Execute(
            "set_print_area",
            workbookPath,
            clearPrintArea: true,
            outputPath: outputPath);
        Assert.Contains("Print area cleared", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintArea));
    }

    [Fact]
    public void SetPrintArea_WithoutRangeOrClear_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_print_area_no_range.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_print_area",
            workbookPath));
        Assert.Contains("Either range or clearPrintArea must be provided", exception.Message);
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
        var result = _tool.Execute(
            "set_print_area",
            workbookPath,
            sheetIndex: 1,
            range: "A1:B5",
            outputPath: outputPath);
        Assert.Contains("sheet 1", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("A1:B5", workbook.Worksheets[1].PageSetup.PrintArea);
    }

    [Fact]
    public void SetPrintArea_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_print_area_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_print_area",
            workbookPath,
            sheetIndex: 99,
            range: "A1:B5"));
    }

    #endregion

    #region SetPrintTitles Tests

    [Fact]
    public void SetPrintTitles_ShouldSetPrintTitles()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_titles.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_print_titles_output.xlsx");
        var result = _tool.Execute(
            "set_print_titles",
            workbookPath,
            rows: "$1:$1",
            columns: "$A:$A",
            outputPath: outputPath);
        Assert.Contains("Print titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$1", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public void SetPrintTitles_OnlyRows_ShouldSetOnlyRows()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_print_titles_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_print_titles_rows_output.xlsx");
        var result = _tool.Execute(
            "set_print_titles",
            workbookPath,
            rows: "$1:$2",
            outputPath: outputPath);
        Assert.Contains("Print titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
    }

    [Fact]
    public void SetPrintTitles_OnlyColumns_ShouldSetOnlyColumns()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_print_titles_cols.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_print_titles_cols_output.xlsx");
        var result = _tool.Execute(
            "set_print_titles",
            workbookPath,
            columns: "$A:$B",
            outputPath: outputPath);
        Assert.Contains("Print titles updated", result);
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
        var result = _tool.Execute(
            "set_print_titles",
            workbookPath,
            clearTitles: true,
            outputPath: outputPath);
        Assert.Contains("Print titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintTitleRows));
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintTitleColumns));
    }

    [Fact]
    public void SetPrintTitles_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_print_titles_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_print_titles",
            workbookPath,
            sheetIndex: 99,
            rows: "$1:$1"));
    }

    #endregion

    #region SetPageSetup Tests

    [Fact]
    public void SetPageSetup_ShouldSetPageSetup()
    {
        var workbookPath = CreateExcelWorkbook("test_set_page_setup.xlsx");
        var outputPath = CreateTestFilePath("test_set_page_setup_output.xlsx");
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            orientation: "Landscape",
            paperSize: "A4",
            outputPath: outputPath);
        Assert.Contains("Page setup updated", result);
        Assert.Contains("orientation=Landscape", result);
        Assert.Contains("paperSize=A4", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
        Assert.Equal(PaperSizeType.PaperA4, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Fact]
    public void SetPageSetup_WithMargins_ShouldSetMargins()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_margins.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_margins_output.xlsx");
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            leftMargin: 0.5,
            rightMargin: 0.5,
            topMargin: 0.75,
            bottomMargin: 0.75,
            outputPath: outputPath);
        Assert.Contains("leftMargin=0.5", result);
        Assert.Contains("rightMargin=0.5", result);
        Assert.Contains("topMargin=0.75", result);
        Assert.Contains("bottomMargin=0.75", result);
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
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            header: "Test Header",
            footer: "Page &P of &N",
            outputPath: outputPath);
        Assert.Contains("header", result);
        Assert.Contains("footer", result);
        using var workbook = new Workbook(outputPath);
        Assert.Contains("Test Header", workbook.Worksheets[0].PageSetup.GetHeader(1));
        Assert.Contains("Page", workbook.Worksheets[0].PageSetup.GetFooter(1));
    }

    [Fact]
    public void SetPageSetup_WithFitToPage_ShouldSetFitToPage()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_fit.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_fit_output.xlsx");
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            fitToPage: true,
            fitToPagesWide: 1,
            fitToPagesTall: 0,
            outputPath: outputPath);
        Assert.Contains("fitToPage", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(1, workbook.Worksheets[0].PageSetup.FitToPagesWide);
        Assert.Equal(0, workbook.Worksheets[0].PageSetup.FitToPagesTall);
    }

    [Fact]
    public void SetPageSetup_AllPaperSizes_ShouldSetCorrectPaperSize()
    {
        var paperSizes = new Dictionary<string, PaperSizeType>
        {
            ["A3"] = PaperSizeType.PaperA3,
            ["A4"] = PaperSizeType.PaperA4,
            ["A5"] = PaperSizeType.PaperA5,
            ["B4"] = PaperSizeType.PaperB4,
            ["B5"] = PaperSizeType.PaperB5,
            ["Letter"] = PaperSizeType.PaperLetter,
            ["Legal"] = PaperSizeType.PaperLegal,
            ["Tabloid"] = PaperSizeType.PaperTabloid,
            ["Executive"] = PaperSizeType.PaperExecutive
        };

        foreach (var (sizeName, expectedType) in paperSizes)
        {
            var workbookPath = CreateExcelWorkbook($"test_paper_{sizeName}.xlsx");
            var outputPath = CreateTestFilePath($"test_paper_{sizeName}_output.xlsx");
            var result = _tool.Execute(
                "set_page_setup",
                workbookPath,
                paperSize: sizeName,
                outputPath: outputPath);
            Assert.Contains($"paperSize={sizeName}", result);
            using var workbook = new Workbook(outputPath);
            Assert.Equal(expectedType, workbook.Worksheets[0].PageSetup.PaperSize);
        }
    }

    [Fact]
    public void SetPageSetup_InvalidPaperSize_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_invalid_paper.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_page_setup",
            workbookPath,
            paperSize: "InvalidSize"));
        Assert.Contains("Invalid paper size", exception.Message);
        Assert.Contains("Supported values", exception.Message);
    }

    [Fact]
    public void SetPageSetup_NoChanges_ShouldReturnNoChangesMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_no_changes.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_no_changes_output.xlsx");
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            outputPath: outputPath);
        Assert.Contains("no changes", result);
    }

    [Fact]
    public void SetPageSetup_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_page_setup_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_page_setup",
            workbookPath,
            sheetIndex: 99,
            orientation: "Landscape"));
    }

    #endregion

    #region SetAll Tests

    [Fact]
    public void SetAll_ShouldSetAllPrintSettings()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_print_settings.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_all_print_settings_output.xlsx");
        var result = _tool.Execute(
            "set_all",
            workbookPath,
            range: "A1:D10",
            orientation: "Portrait",
            paperSize: "A4",
            outputPath: outputPath);
        Assert.Contains("Print settings updated", result);
        Assert.Contains("printArea=A1:D10", result);
        Assert.Contains("orientation=Portrait", result);
        Assert.Contains("paperSize=A4", result);
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
        var result = _tool.Execute(
            "set_all",
            workbookPath,
            range: "A1:J20",
            rows: "$1:$2",
            columns: "$A:$A",
            outputPath: outputPath);
        Assert.Contains("printArea=A1:J20", result);
        Assert.Contains("printTitleRows=$1:$2", result);
        Assert.Contains("printTitleColumns=$A:$A", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public void SetAll_WithFitToPage_ShouldSetFitToPage()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_fit.xlsx", 50, 20);
        var outputPath = CreateTestFilePath("test_set_all_fit_output.xlsx");
        var result = _tool.Execute(
            "set_all",
            workbookPath,
            range: "A1:T50",
            fitToPage: true,
            fitToPagesWide: 1,
            fitToPagesTall: 0,
            outputPath: outputPath);
        Assert.Contains("fitToPage", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(1, workbook.Worksheets[0].PageSetup.FitToPagesWide);
        Assert.Equal(0, workbook.Worksheets[0].PageSetup.FitToPagesTall);
    }

    [Fact]
    public void SetAll_WithAllOptions_ShouldSetAllOptions()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_options.xlsx", 30, 15);
        var outputPath = CreateTestFilePath("test_set_all_options_output.xlsx");
        var result = _tool.Execute(
            "set_all",
            workbookPath,
            range: "A1:O30",
            rows: "$1:$1",
            columns: "$A:$A",
            orientation: "Landscape",
            paperSize: "Letter",
            leftMargin: 0.5,
            rightMargin: 0.5,
            topMargin: 0.75,
            bottomMargin: 0.75,
            header: "Report Title",
            footer: "Page &P",
            outputPath: outputPath);
        Assert.Contains("printArea=A1:O30", result);
        Assert.Contains("printTitleRows=$1:$1", result);
        Assert.Contains("orientation=Landscape", result);
        Assert.Contains("paperSize=Letter", result);
        Assert.Contains("leftMargin=0.5", result);
        Assert.Contains("header", result);
        Assert.Contains("footer", result);

        using var workbook = new Workbook(outputPath);
        var pageSetup = workbook.Worksheets[0].PageSetup;
        Assert.Equal("A1:O30", pageSetup.PrintArea);
        Assert.Equal(PageOrientationType.Landscape, pageSetup.Orientation);
        Assert.Equal(PaperSizeType.PaperLetter, pageSetup.PaperSize);
        Assert.Equal(0.5, pageSetup.LeftMargin, 5);
    }

    [Fact]
    public void SetAll_NoChanges_ShouldReturnNoChangesMessage()
    {
        var workbookPath = CreateExcelWorkbook("test_set_all_no_changes.xlsx");
        var outputPath = CreateTestFilePath("test_set_all_no_changes_output.xlsx");
        var result = _tool.Execute(
            "set_all",
            workbookPath,
            outputPath: outputPath);
        Assert.Contains("no changes", result);
    }

    [Fact]
    public void SetAll_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_set_all_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "set_all",
            workbookPath,
            sheetIndex: 99,
            range: "A1:D10"));
    }

    #endregion

    [Fact]
    public void ExecuteAsync_DefaultOutputPath_ShouldUseInputPath()
    {
        var workbookPath = CreateExcelWorkbook("test_default_output.xlsx");
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            orientation: "Landscape");
        Assert.Contains(workbookPath, result);
        using var workbook = new Workbook(workbookPath);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
    }

    [Fact]
    public void SetPageSetup_CaseInsensitivePaperSize_ShouldWork()
    {
        var workbookPath = CreateExcelWorkbook("test_case_insensitive.xlsx");
        var outputPath = CreateTestFilePath("test_case_insensitive_output.xlsx");
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            paperSize: "letter", // lowercase
            outputPath: outputPath);
        Assert.Contains("paperSize=letter", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PaperSizeType.PaperLetter, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Fact]
    public void SetPageSetup_CaseInsensitiveOrientation_ShouldWork()
    {
        var workbookPath = CreateExcelWorkbook("test_case_orientation.xlsx");
        var outputPath = CreateTestFilePath("test_case_orientation_output.xlsx");
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            orientation: "landscape", // lowercase
            outputPath: outputPath);
        Assert.Contains("orientation=landscape", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_operation.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid_operation",
            workbookPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void SetPageSetup_InvalidOrientation_ShouldDefaultToPortrait()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_orientation.xlsx");
        var outputPath = CreateTestFilePath("test_invalid_orientation_output.xlsx");

        // Act - Invalid orientation should default to Portrait
        var result = _tool.Execute(
            "set_page_setup",
            workbookPath,
            outputPath: outputPath,
            orientation: "Diagonal");

        // Assert - Should succeed and use Portrait as default
        Assert.Contains("orientation", result, StringComparison.OrdinalIgnoreCase);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PageOrientationType.Portrait, workbook.Worksheets[0].PageSetup.Orientation);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void SetPrintArea_WithSessionId_ShouldSetInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_print_area.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "set_print_area",
            sessionId: sessionId,
            range: "A1:D10");
        Assert.Contains("Print area set to A1:D10", result);

        // Verify in-memory workbook has the print area
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    [Fact]
    public void SetPageSetup_WithSessionId_ShouldSetInMemory()
    {
        var workbookPath = CreateExcelWorkbook("test_session_page_setup.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "set_page_setup",
            sessionId: sessionId,
            orientation: "Landscape",
            paperSize: "A4");
        Assert.Contains("Page setup updated", result);

        // Verify in-memory workbook has the settings
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
        Assert.Equal(PaperSizeType.PaperA4, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Fact]
    public void SetPrintTitles_WithSessionId_ShouldSetInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_print_titles.xlsx", 10, 5);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "set_print_titles",
            sessionId: sessionId,
            rows: "$1:$2",
            columns: "$A:$A");
        Assert.Contains("Print titles updated", result);

        // Verify in-memory workbook has the print titles
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("set_print_area", sessionId: "invalid_session_id", range: "A1:D10"));
    }

    #endregion
}