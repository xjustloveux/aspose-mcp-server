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

    #region General

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
    public void FormatCells_WithColors_ShouldApplyColors()
    {
        var workbookPath = CreateExcelWorkbook("test_format_colors.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_colors_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", fontColor: "#FF0000",
            backgroundColor: "#FFFF00", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        var fontColor = style.Font.Color.ToArgb() & 0xFFFFFF;
        var fgColor = style.ForegroundColor.ToArgb() & 0xFFFFFF;
        Assert.Equal(0xFF0000, fontColor);
        Assert.Equal(0xFFFF00, fgColor);
    }

    [Fact]
    public void FormatCells_WithAlignment_ShouldApplyAlignment()
    {
        var workbookPath = CreateExcelWorkbook("test_format_alignment.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_alignment_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", horizontalAlignment: "Center",
            verticalAlignment: "Center", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
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
        _tool.Execute("format", workbookPath, range: "A1", borderStyle: "Thin",
            borderColor: "#000000", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        var hasBorder = style.Borders[BorderType.TopBorder].LineStyle != CellBorderType.None ||
                        style.Borders[BorderType.BottomBorder].LineStyle != CellBorderType.None;
        Assert.True(hasBorder);
    }

    [Fact]
    public void FormatCells_WithDoubleBorder_ShouldApplyCorrectStyle()
    {
        var workbookPath = CreateExcelWorkbook("test_format_border_double.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_border_double_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", borderStyle: "double", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(CellBorderType.Double, style.Borders[BorderType.TopBorder].LineStyle);
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
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void FormatCells_WithBuiltInNumberFormat_ShouldApplyFormat()
    {
        var workbookPath = CreateExcelWorkbook("test_format_builtin.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 1234.56;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_builtin_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", numberFormat: "4", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(4, style.Number);
    }

    [Fact]
    public void FormatCells_WithPatternFill_ShouldApplyPattern()
    {
        var workbookPath = CreateExcelWorkbook("test_format_pattern.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Pattern Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_pattern_output.xlsx");
        _tool.Execute("format", workbookPath, range: "A1", patternType: "DiagonalStripe",
            backgroundColor: "#FF0000", patternColor: "#FFFFFF", outputPath: outputPath);
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
        _tool.Execute("format", workbookPath, range: "A1", patternType: "Gray50",
            backgroundColor: "#0000FF", outputPath: outputPath);
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(BackgroundType.Gray50, style.Pattern);
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
    public void GetFormat_WithCellParameter_ShouldReturnFormatInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_cell.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var result = _tool.Execute("get_format", workbookPath, cell: "A1");
        Assert.Contains("A1", result);
        Assert.Contains("fontName", result);
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
        Assert.Contains("A1", result);
        Assert.Contains("A2", result);
        Assert.Contains("\"count\": 2", result);
    }

    [Fact]
    public void GetFormat_WithFieldsParameter_ShouldReturnOnlyRequestedFields()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_fields.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "font");
        Assert.Contains("fontName", result);
        Assert.Contains("fontSize", result);
        Assert.DoesNotContain("borders", result);
    }

    [Fact]
    public void GetFormat_WithMultipleFields_ShouldReturnRequestedFields()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_multi_fields.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "font,color");
        Assert.Contains("fontName", result);
        Assert.Contains("fontColor", result);
        Assert.DoesNotContain("borders", result);
    }

    [Fact]
    public void GetFormat_WithValueField_ShouldReturnValueInfo()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_value.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "TestValue";
        workbook.Save(workbookPath);

        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "value");
        Assert.Contains("TestValue", result);
        Assert.Contains("dataType", result);
        Assert.DoesNotContain("fontName", result);
    }

    [Fact]
    public void GetFormat_WithColorField_ShouldIncludePatternType()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_color.xlsx");
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
        Assert.Contains("patternType", result);
        Assert.Contains("foregroundColor", result);
        Assert.Contains("backgroundColor", result);
    }

    [Fact]
    public void GetFormat_WithAlignmentField_ShouldReturnAlignmentOnly()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_align.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var result = _tool.Execute("get_format", workbookPath, range: "A1", fields: "alignment");
        Assert.Contains("horizontalAlignment", result);
        Assert.Contains("verticalAlignment", result);
        Assert.DoesNotContain("fontName", result);
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

    [Fact]
    public void CopySheetFormat_WithColumnWidthsOnly_ShouldCopyColumnWidths()
    {
        var workbookPath = CreateExcelWorkbook("test_copy_col_widths.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells.SetColumnWidth(0, 20);
        workbook.Worksheets.Add("TargetSheet");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_col_widths_output.xlsx");
        var result = _tool.Execute("copy_sheet_format", workbookPath, sourceSheetIndex: 0,
            targetSheetIndex: 1, copyColumnWidths: true, copyRowHeights: false, outputPath: outputPath);
        Assert.Contains("Sheet format copied", result);
    }

    [Theory]
    [InlineData("FORMAT")]
    [InlineData("Format")]
    [InlineData("format")]
    public void Operation_ShouldBeCaseInsensitive_Format(string operation)
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

    [Theory]
    [InlineData("GET_FORMAT")]
    [InlineData("Get_Format")]
    [InlineData("get_format")]
    public void Operation_ShouldBeCaseInsensitive_GetFormat(string operation)
    {
        var workbookPath = CreateExcelWorkbook($"test_case_get_{operation.Replace("_", "")}.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var result = _tool.Execute(operation, workbookPath, range: "A1");
        Assert.Contains("A1", result);
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
    public void Format_WithMissingRangeAndRanges_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_format_no_range.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", workbookPath, bold: true));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Format_WithInvalidColor_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_format_invalid_color.xlsx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("format", workbookPath, range: "A1", backgroundColor: "invalid_color"));
    }

    [Fact]
    public void GetFormat_WithMissingCellAndRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_no_cell.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_format", workbookPath));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetFormat_WithInvalidRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_format_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get_format", workbookPath, range: "INVALID"));
        Assert.Contains("Invalid", ex.Message);
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get_format", range: "A1"));
    }

    #endregion

    #region Session

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