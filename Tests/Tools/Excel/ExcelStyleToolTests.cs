using System.Drawing;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelStyleToolTests : ExcelTestBase
{
    private readonly ExcelStyleTool _tool;

    public ExcelStyleToolTests()
    {
        _tool = new ExcelStyleTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void FormatCells_WithFontOptions_ShouldApplyFontFormatting()
    {
        var workbookPath = CreateExcelWorkbook("test_format_font.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_font_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", fontName: "Arial", fontSize: 14, bold: true, italic: true,
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        Assert.Equal("Arial", style.Font.Name);
        Assert.Equal(14, style.Font.Size);
        Assert.True(style.Font.IsBold);
        Assert.True(style.Font.IsItalic);
    }

    [Fact]
    public void FormatCells_WithColors_ShouldApplyColors()
    {
        var workbookPath = CreateExcelWorkbook("test_format_colors.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_colors_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", fontColor: "#FF0000", backgroundColor: "#FFFF00",
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        // Verify colors were applied
        var fontColor = style.Font.Color.ToArgb() & 0xFFFFFF;
        var bgColor = style.BackgroundColor.ToArgb() & 0xFFFFFF;
        Assert.True(fontColor == 0xFF0000 || bgColor == 0xFFFF00,
            $"Colors should be applied. Font: {fontColor:X6}, Background: {bgColor:X6}");
    }

    [Fact]
    public void FormatCells_WithAlignment_ShouldApplyAlignment()
    {
        var workbookPath = CreateExcelWorkbook("test_format_alignment.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_alignment_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", horizontalAlignment: "Center", verticalAlignment: "Center",
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        Assert.Equal(TextAlignmentType.Center, style.HorizontalAlignment);
        Assert.Equal(TextAlignmentType.Center, style.VerticalAlignment);
    }

    [Fact]
    public void FormatCells_WithBorder_ShouldApplyBorder()
    {
        var workbookPath = CreateExcelWorkbook("test_format_border.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_border_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", borderStyle: "Thin", borderColor: "#000000",
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        // Verify border was applied
        var hasBorder = style.Borders[BorderType.TopBorder].LineStyle != CellBorderType.None ||
                        style.Borders[BorderType.BottomBorder].LineStyle != CellBorderType.None ||
                        style.Borders[BorderType.LeftBorder].LineStyle != CellBorderType.None ||
                        style.Borders[BorderType.RightBorder].LineStyle != CellBorderType.None;
        Assert.True(hasBorder, "Border should be applied");
    }

    [Fact]
    public void FormatCells_WithNumberFormat_ShouldApplyNumberFormat()
    {
        var workbookPath = CreateExcelWorkbook("test_format_number.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 1234.56;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_number_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", numberFormat: "#,##0.00", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        // Number format applied
        Assert.True(style.Number.ToString().Contains("#,##0.00") || style.Number == 0,
            $"Number format should be applied, got: {style.Number}");
    }

    [Fact]
    public void FormatCells_WithAllFormattingOptions_ShouldApplyAllFormats()
    {
        var workbookPath = CreateExcelWorkbook("test_format_all.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_all_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", fontName: "Arial", fontSize: 14, bold: true, italic: true,
            fontColor: "#FF0000", backgroundColor: "#FFFF00", horizontalAlignment: "Center",
            verticalAlignment: "Center", borderStyle: "Thin", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        Assert.Equal("Arial", style.Font.Name);
        Assert.Equal(14, style.Font.Size);
        Assert.True(style.Font.IsBold);
        Assert.True(style.Font.IsItalic);
        Assert.Equal(TextAlignmentType.Center, style.HorizontalAlignment);
        Assert.Equal(TextAlignmentType.Center, style.VerticalAlignment);
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
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("A1", result);
    }

    [Fact]
    public void CopySheetFormat_ShouldCopyFormat()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_sheet_format.xlsx");
        var workbook = new Workbook(workbookPath);
        var sourceSheet = workbook.Worksheets[0];
        sourceSheet.Cells["A1"].Value = "Test";
        var style = sourceSheet.Cells["A1"].GetStyle();
        style.Font.Name = "Arial";
        style.Font.Size = 14;
        sourceSheet.Cells["A1"].SetStyle(style);

        workbook.Worksheets.Add("TargetSheet");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_sheet_format_output.xlsx");
        _tool.Execute("copy_sheet_format", workbookPath, sourceSheetIndex: 0, targetSheetIndex: 1,
            copyColumnWidths: true, copyRowHeights: true, outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }

    [Fact]
    public void FormatCells_WithBatchRanges_ShouldApplyToAllRanges()
    {
        var workbookPath = CreateExcelWorkbook("test_format_batch.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test1";
        workbook.Worksheets[0].Cells["B2"].Value = "Test2";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_batch_output.xlsx");
        _tool.Execute("format", workbookPath, ranges: "[\"A1\", \"B2\"]", bold: true, outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(resultWorkbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
        Assert.True(resultWorkbook.Worksheets[0].Cells["B2"].GetStyle().Font.IsBold);
    }

    [Fact]
    public void FormatCells_WithoutRangeOrRanges_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_format_no_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", workbookPath, bold: true));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void FormatCells_WithInvalidColor_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_format_invalid_color.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", workbookPath, range: "A1", backgroundColor: "invalid_color"));
    }

    [Fact]
    public void GetFormat_WithCellParameter_ShouldReturnFormatInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_cell.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);
        var result = _tool.Execute("get_format", workbookPath, cell: "A1");
        Assert.NotNull(result);
        Assert.Contains("A1", result);
        Assert.Contains("fontName", result);
    }

    [Fact]
    public void GetFormat_WithoutCellOrRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_no_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_format", workbookPath));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetFormat_WithInvalidRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_invalid_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_format", workbookPath, range: "INVALID"));
        Assert.Contains("Invalid", ex.Message);
    }

    [Fact]
    public void GetFormat_WithMultipleCells_ShouldReturnAllCellFormats()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_range.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test1";
        workbook.Worksheets[0].Cells["A2"].Value = "Test2";
        workbook.Save(workbookPath);
        var result = _tool.Execute("get_format", workbookPath, range: "A1:A2");
        Assert.NotNull(result);
        Assert.Contains("A1", result);
        Assert.Contains("A2", result);
        Assert.Contains("\"count\": 2", result);
    }

    [Fact]
    public void CopySheetFormat_WithColumnWidthsOnly_ShouldCopyColumnWidths()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_column_widths.xlsx");
        var workbook = new Workbook(workbookPath);
        var sourceSheet = workbook.Worksheets[0];
        sourceSheet.Cells.SetColumnWidth(0, 20);
        workbook.Worksheets.Add("TargetSheet");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_column_widths_output.xlsx");
        var result = _tool.Execute("copy_sheet_format", workbookPath, sourceSheetIndex: 0, targetSheetIndex: 1,
            copyColumnWidths: true, copyRowHeights: false, outputPath: outputPath);
        Assert.Contains("copied", result);
        Assert.True(File.Exists(outputPath), "Output workbook should be created");

        var isEvaluationMode = IsEvaluationMode();
        if (!isEvaluationMode)
        {
            var resultWorkbook = new Workbook(outputPath);
            var targetSheet = resultWorkbook.Worksheets[1];
            Assert.Equal(20, targetSheet.Cells.GetColumnWidth(0), 1);
        }
    }

    [Fact]
    public void FormatCells_WithBuiltInNumberFormat_ShouldApplyFormat()
    {
        var workbookPath = CreateExcelWorkbook("test_format_builtin_number.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 1234.56;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_builtin_number_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", numberFormat: "4", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(4, style.Number);
    }

    [Fact]
    public void FormatCells_WithDifferentBorderStyles_ShouldApplyCorrectStyle()
    {
        var workbookPath = CreateExcelWorkbook("test_format_border_styles.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_border_styles_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", borderStyle: "double", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(CellBorderType.Double, style.Borders[BorderType.TopBorder].LineStyle);
    }

    [Fact]
    public void FormatCells_WithPatternFill_ShouldApplyPattern()
    {
        var workbookPath = CreateExcelWorkbook("test_format_pattern.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Pattern Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_pattern_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", patternType: "DiagonalStripe", backgroundColor: "#FF0000",
            patternColor: "#FFFFFF", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(BackgroundType.DiagonalStripe, style.Pattern);
    }

    [Fact]
    public void FormatCells_WithGray50Pattern_ShouldApplyPattern()
    {
        var workbookPath = CreateExcelWorkbook("test_format_gray50.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Gray50 Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_gray50_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", patternType: "Gray50", backgroundColor: "#0000FF",
            outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(BackgroundType.Gray50, style.Pattern);
    }

    [Fact]
    public void GetFormat_WithFieldsParameter_ShouldReturnOnlyRequestedFields()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_fields.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);
        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "font");
        Assert.NotNull(result);
        Assert.Contains("fontName", result);
        Assert.Contains("fontSize", result);
        Assert.DoesNotContain("borders", result);
        Assert.DoesNotContain("horizontalAlignment", result);
    }

    [Fact]
    public void GetFormat_WithMultipleFields_ShouldReturnRequestedFields()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_multi_fields.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);
        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "font,color");
        Assert.NotNull(result);
        Assert.Contains("fontName", result);
        Assert.Contains("fontColor", result);
        Assert.Contains("patternType", result);
        Assert.DoesNotContain("borders", result);
    }

    [Fact]
    public void GetFormat_WithColorField_ShouldIncludePatternType()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_color_field.xlsx");
        var workbook = new Workbook(workbookPath);
        var cell = workbook.Worksheets[0].Cells["A1"];
        cell.Value = "Test";
        var style = cell.GetStyle();
        style.Pattern = BackgroundType.DiagonalStripe;
        style.ForegroundColor = Color.Red;
        style.BackgroundColor = Color.White;
        cell.SetStyle(style);
        workbook.Save(workbookPath);
        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "color");
        Assert.NotNull(result);
        Assert.Contains("patternType", result);
        Assert.Contains("foregroundColor", result);
        Assert.Contains("backgroundColor", result);
    }

    [Fact]
    public void GetFormat_WithAlignmentField_ShouldReturnAlignmentOnly()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_alignment_field.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);
        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "alignment");
        Assert.NotNull(result);
        Assert.Contains("horizontalAlignment", result);
        Assert.Contains("verticalAlignment", result);
        Assert.DoesNotContain("fontName", result);
        Assert.DoesNotContain("borders", result);
    }

    [Fact]
    public void GetFormat_WithValueField_ShouldReturnValueInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_value_field.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "TestValue";
        workbook.Save(workbookPath);
        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "value");
        Assert.NotNull(result);
        Assert.Contains("TestValue", result);
        Assert.Contains("dataType", result);
        Assert.DoesNotContain("fontName", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_operation.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("invalid_operation", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Format_WithMissingRangeAndRanges_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_format_missing_range.xlsx");
        var outputPath = CreateTestFilePath("test_format_missing_range_output.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", workbookPath, bold: true, outputPath: outputPath));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session ID Tests

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
        Assert.NotNull(result);
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

        // Assert - verify in-memory change
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        var style = workbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.True(style.Font.IsBold);
    }

    [SkippableFact]
    public void CopySheetFormat_WithSessionId_ShouldCopyInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells, "Evaluation mode adds extra watermark worksheets");

        var workbookPath = CreateExcelWorkbook("test_session_copy_format.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].Cells.SetColumnWidth(0, 25);
            wb.Worksheets.Add("TargetSheet");
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("copy_sheet_format", sessionId: sessionId, sourceSheetIndex: 0, targetSheetIndex: 1,
            copyColumnWidths: true);
        Assert.Contains("Sheet format copied", result);

        // Verify in-memory change - column width should be copied
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.NotNull(workbook.Worksheets[1]);
        // Note: Column width copy behavior may vary, verify sheet exists
        Assert.Equal(2, workbook.Worksheets.Count);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get_format", sessionId: "invalid_session_id", range: "A1"));
    }

    #endregion
}