using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelPrintSettingsToolTests : ExcelTestBase
{
    private readonly ExcelPrintSettingsTool _tool = new();

    #region SetPrintArea Tests

    [Fact]
    public async Task SetPrintArea_ShouldSetPrintArea()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_area.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_print_area_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_area",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:D10"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Print area set to A1:D10", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    [Fact]
    public async Task SetPrintArea_MultipleRanges_ShouldSetMultiplePrintAreas()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_area_multi.xlsx", 10, 10);
        var outputPath = CreateTestFilePath("test_set_print_area_multi_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_area",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:D10,F1:H10"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Print area set to A1:D10,F1:H10", result);
        using var workbook = new Workbook(outputPath);
        Assert.Contains("A1:D10", workbook.Worksheets[0].PageSetup.PrintArea);
        Assert.Contains("F1:H10", workbook.Worksheets[0].PageSetup.PrintArea);
    }

    [Fact]
    public async Task SetPrintArea_ClearPrintArea_ShouldClearPrintArea()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_clear_print_area.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].PageSetup.PrintArea = "A1:D10";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_clear_print_area_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_area",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["clearPrintArea"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Print area cleared", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintArea));
    }

    [Fact]
    public async Task SetPrintArea_WithoutRangeOrClear_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_print_area_no_range.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_area",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Either range or clearPrintArea must be provided", exception.Message);
    }

    [Fact]
    public async Task SetPrintArea_WithSheetIndex_ShouldSetPrintAreaOnCorrectSheet()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_print_area_sheet.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets.Add("Sheet2");
            wb.Worksheets[1].Cells["A1"].PutValue("Test");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_print_area_sheet_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_area",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sheetIndex"] = 1,
            ["range"] = "A1:B5"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("sheet 1", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("A1:B5", workbook.Worksheets[1].PageSetup.PrintArea);
    }

    [Fact]
    public async Task SetPrintArea_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_print_area_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_area",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["range"] = "A1:B5"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region SetPrintTitles Tests

    [Fact]
    public async Task SetPrintTitles_ShouldSetPrintTitles()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_print_titles.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_print_titles_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_titles",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["rows"] = "$1:$1",
            ["columns"] = "$A:$A"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Print titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$1", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public async Task SetPrintTitles_OnlyRows_ShouldSetOnlyRows()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_print_titles_rows.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_print_titles_rows_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_titles",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["rows"] = "$1:$2"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Print titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
    }

    [Fact]
    public async Task SetPrintTitles_OnlyColumns_ShouldSetOnlyColumns()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_print_titles_cols.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_print_titles_cols_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_titles",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["columns"] = "$A:$B"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Print titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$A:$B", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public async Task SetPrintTitles_ClearTitles_ShouldClearPrintTitles()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_clear_print_titles.xlsx", 10, 5);
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].PageSetup.PrintTitleRows = "$1:$1";
            wb.Worksheets[0].PageSetup.PrintTitleColumns = "$A:$A";
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_clear_print_titles_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_titles",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["clearTitles"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Print titles updated", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintTitleRows));
        Assert.True(string.IsNullOrEmpty(workbook.Worksheets[0].PageSetup.PrintTitleColumns));
    }

    [Fact]
    public async Task SetPrintTitles_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_print_titles_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_print_titles",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["rows"] = "$1:$1"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region SetPageSetup Tests

    [Fact]
    public async Task SetPageSetup_ShouldSetPageSetup()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_page_setup.xlsx");
        var outputPath = CreateTestFilePath("test_set_page_setup_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["orientation"] = "Landscape",
            ["paperSize"] = "A4"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Page setup updated", result);
        Assert.Contains("orientation=Landscape", result);
        Assert.Contains("paperSize=A4", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
        Assert.Equal(PaperSizeType.PaperA4, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Fact]
    public async Task SetPageSetup_WithMargins_ShouldSetMargins()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_page_setup_margins.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_margins_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["leftMargin"] = 0.5,
            ["rightMargin"] = 0.5,
            ["topMargin"] = 0.75,
            ["bottomMargin"] = 0.75
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task SetPageSetup_WithHeaderFooter_ShouldSetHeaderFooter()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_page_setup_header_footer.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_header_footer_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["header"] = "Test Header",
            ["footer"] = "Page &P of &N"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("header", result);
        Assert.Contains("footer", result);
        using var workbook = new Workbook(outputPath);
        Assert.Contains("Test Header", workbook.Worksheets[0].PageSetup.GetHeader(1));
        Assert.Contains("Page", workbook.Worksheets[0].PageSetup.GetFooter(1));
    }

    [Fact]
    public async Task SetPageSetup_WithFitToPage_ShouldSetFitToPage()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_page_setup_fit.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_fit_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["fitToPage"] = true,
            ["fitToPagesWide"] = 1,
            ["fitToPagesTall"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("fitToPage", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(1, workbook.Worksheets[0].PageSetup.FitToPagesWide);
        Assert.Equal(0, workbook.Worksheets[0].PageSetup.FitToPagesTall);
    }

    [Fact]
    public async Task SetPageSetup_AllPaperSizes_ShouldSetCorrectPaperSize()
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
            // Arrange
            var workbookPath = CreateExcelWorkbook($"test_paper_{sizeName}.xlsx");
            var outputPath = CreateTestFilePath($"test_paper_{sizeName}_output.xlsx");
            var arguments = new JsonObject
            {
                ["operation"] = "set_page_setup",
                ["path"] = workbookPath,
                ["outputPath"] = outputPath,
                ["paperSize"] = sizeName
            };

            // Act
            var result = await _tool.ExecuteAsync(arguments);

            // Assert
            Assert.Contains($"paperSize={sizeName}", result);
            using var workbook = new Workbook(outputPath);
            Assert.Equal(expectedType, workbook.Worksheets[0].PageSetup.PaperSize);
        }
    }

    [Fact]
    public async Task SetPageSetup_InvalidPaperSize_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_page_setup_invalid_paper.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["paperSize"] = "InvalidSize"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid paper size", exception.Message);
        Assert.Contains("Supported values", exception.Message);
    }

    [Fact]
    public async Task SetPageSetup_NoChanges_ShouldReturnNoChangesMessage()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_page_setup_no_changes.xlsx");
        var outputPath = CreateTestFilePath("test_page_setup_no_changes_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("no changes", result);
    }

    [Fact]
    public async Task SetPageSetup_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_page_setup_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["orientation"] = "Landscape"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region SetAll Tests

    [Fact]
    public async Task SetAll_ShouldSetAllPrintSettings()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_print_settings.xlsx", 10, 5);
        var outputPath = CreateTestFilePath("test_set_all_print_settings_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_all",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:D10",
            ["orientation"] = "Portrait",
            ["paperSize"] = "A4"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task SetAll_WithPrintTitles_ShouldSetPrintTitles()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_titles.xlsx", 20, 10);
        var outputPath = CreateTestFilePath("test_set_all_titles_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_all",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:J20",
            ["rows"] = "$1:$2",
            ["columns"] = "$A:$A"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("printArea=A1:J20", result);
        Assert.Contains("printTitleRows=$1:$2", result);
        Assert.Contains("printTitleColumns=$A:$A", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("$1:$2", workbook.Worksheets[0].PageSetup.PrintTitleRows);
        Assert.Equal("$A:$A", workbook.Worksheets[0].PageSetup.PrintTitleColumns);
    }

    [Fact]
    public async Task SetAll_WithFitToPage_ShouldSetFitToPage()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_fit.xlsx", 50, 20);
        var outputPath = CreateTestFilePath("test_set_all_fit_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_all",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:T50",
            ["fitToPage"] = true,
            ["fitToPagesWide"] = 1,
            ["fitToPagesTall"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("fitToPage", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(1, workbook.Worksheets[0].PageSetup.FitToPagesWide);
        Assert.Equal(0, workbook.Worksheets[0].PageSetup.FitToPagesTall);
    }

    [Fact]
    public async Task SetAll_WithAllOptions_ShouldSetAllOptions()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_set_all_options.xlsx", 30, 15);
        var outputPath = CreateTestFilePath("test_set_all_options_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_all",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1:O30",
            ["rows"] = "$1:$1",
            ["columns"] = "$A:$A",
            ["orientation"] = "Landscape",
            ["paperSize"] = "Letter",
            ["leftMargin"] = 0.5,
            ["rightMargin"] = 0.5,
            ["topMargin"] = 0.75,
            ["bottomMargin"] = 0.75,
            ["header"] = "Report Title",
            ["footer"] = "Page &P"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task SetAll_NoChanges_ShouldReturnNoChangesMessage()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_all_no_changes.xlsx");
        var outputPath = CreateTestFilePath("test_set_all_no_changes_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_all",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("no changes", result);
    }

    [Fact]
    public async Task SetAll_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_set_all_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_all",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99,
            ["range"] = "A1:D10"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region General Tests

    [Fact]
    public async Task ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_operation.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "invalid_operation",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task ExecuteAsync_DefaultOutputPath_ShouldUseInputPath()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_default_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["orientation"] = "Landscape"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains(workbookPath, result);
        using var workbook = new Workbook(workbookPath);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
    }

    [Fact]
    public async Task SetPageSetup_CaseInsensitivePaperSize_ShouldWork()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_case_insensitive.xlsx");
        var outputPath = CreateTestFilePath("test_case_insensitive_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["paperSize"] = "letter" // lowercase
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("paperSize=letter", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PaperSizeType.PaperLetter, workbook.Worksheets[0].PageSetup.PaperSize);
    }

    [Fact]
    public async Task SetPageSetup_CaseInsensitiveOrientation_ShouldWork()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_case_orientation.xlsx");
        var outputPath = CreateTestFilePath("test_case_orientation_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_page_setup",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["orientation"] = "landscape" // lowercase
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("orientation=landscape", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PageOrientationType.Landscape, workbook.Worksheets[0].PageSetup.Orientation);
    }

    #endregion
}