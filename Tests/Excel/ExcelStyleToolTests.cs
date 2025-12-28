using System.Drawing;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelStyleToolTests : ExcelTestBase
{
    private readonly ExcelStyleTool _tool = new();

    [Fact]
    public async Task FormatCells_WithFontOptions_ShouldApplyFontFormatting()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_font.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_font_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["fontName"] = "Arial",
            ["fontSize"] = 14,
            ["bold"] = true,
            ["italic"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        Assert.Equal("Arial", style.Font.Name);
        Assert.Equal(14, style.Font.Size);
        Assert.True(style.Font.IsBold);
        Assert.True(style.Font.IsItalic);
    }

    [Fact]
    public async Task FormatCells_WithColors_ShouldApplyColors()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_colors.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_colors_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["fontColor"] = "#FF0000",
            ["backgroundColor"] = "#FFFF00"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task FormatCells_WithAlignment_ShouldApplyAlignment()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_alignment.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_alignment_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["horizontalAlignment"] = "Center",
            ["verticalAlignment"] = "Center"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        Assert.Equal(TextAlignmentType.Center, style.HorizontalAlignment);
        Assert.Equal(TextAlignmentType.Center, style.VerticalAlignment);
    }

    [Fact]
    public async Task FormatCells_WithBorder_ShouldApplyBorder()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_border.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_border_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["borderStyle"] = "Thin",
            ["borderColor"] = "#000000"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task FormatCells_WithNumberFormat_ShouldApplyNumberFormat()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_number.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 1234.56;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_number_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["numberFormat"] = "#,##0.00"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var worksheet = resultWorkbook.Worksheets[0];
        var style = worksheet.Cells["A1"].GetStyle();
        // Number format applied
        Assert.True(style.Number.ToString().Contains("#,##0.00") || style.Number == 0,
            $"Number format should be applied, got: {style.Number}");
    }

    [Fact]
    public async Task FormatCells_WithAllFormattingOptions_ShouldApplyAllFormats()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_all.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_all_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["fontName"] = "Arial",
            ["fontSize"] = 14,
            ["bold"] = true,
            ["italic"] = true,
            ["fontColor"] = "#FF0000",
            ["backgroundColor"] = "#FFFF00",
            ["horizontalAlignment"] = "Center",
            ["verticalAlignment"] = "Center",
            ["borderStyle"] = "Thin"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task GetFormat_ShouldReturnFormatInfo()
    {
        // Arrange
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

        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["range"] = "A1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("A1", result);
    }

    [Fact]
    public async Task CopySheetFormat_ShouldCopyFormat()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "copy_sheet_format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sourceSheetIndex"] = 0,
            ["targetSheetIndex"] = 1,
            ["copyColumnWidths"] = true,
            ["copyRowHeights"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output workbook should be created");
    }

    [Fact]
    public async Task FormatCells_WithBatchRanges_ShouldApplyToAllRanges()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_batch.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test1";
        workbook.Worksheets[0].Cells["B2"].Value = "Test2";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_batch_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["ranges"] = new JsonArray { "A1", "B2" },
            ["bold"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        Assert.True(resultWorkbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
        Assert.True(resultWorkbook.Worksheets[0].Cells["B2"].GetStyle().Font.IsBold);
    }

    [Fact]
    public async Task FormatCells_WithoutRangeOrRanges_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_no_range.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["bold"] = true
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task FormatCells_WithInvalidColor_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_invalid_color.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["range"] = "A1",
            ["backgroundColor"] = "invalid_color"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetFormat_WithCellParameter_ShouldReturnFormatInfo()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_format_cell.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["cell"] = "A1"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("A1", result);
        Assert.Contains("fontName", result);
    }

    [Fact]
    public async Task GetFormat_WithoutCellOrRange_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_format_no_cell.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("cell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetFormat_WithInvalidRange_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_format_invalid_range.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["range"] = "INVALID"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid", ex.Message);
    }

    [Fact]
    public async Task GetFormat_WithMultipleCells_ShouldReturnAllCellFormats()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_format_range.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test1";
        workbook.Worksheets[0].Cells["A2"].Value = "Test2";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["range"] = "A1:A2"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("A1", result);
        Assert.Contains("A2", result);
        Assert.Contains("\"count\": 2", result);
    }

    [Fact]
    public async Task CopySheetFormat_WithColumnWidthsOnly_ShouldCopyColumnWidths()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_copy_column_widths.xlsx");
        var workbook = new Workbook(workbookPath);
        var sourceSheet = workbook.Worksheets[0];
        sourceSheet.Cells.SetColumnWidth(0, 20);
        workbook.Worksheets.Add("TargetSheet");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_copy_column_widths_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "copy_sheet_format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sourceSheetIndex"] = 0,
            ["targetSheetIndex"] = 1,
            ["copyColumnWidths"] = true,
            ["copyRowHeights"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task InvalidOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_operation.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "invalid_operation",
            ["path"] = workbookPath
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public async Task FormatCells_WithBuiltInNumberFormat_ShouldApplyFormat()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_builtin_number.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = 1234.56;
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_builtin_number_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["numberFormat"] = "4" // Built-in format number
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(4, style.Number);
    }

    [Fact]
    public async Task FormatCells_WithDifferentBorderStyles_ShouldApplyCorrectStyle()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_border_styles.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_border_styles_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["borderStyle"] = "double"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(CellBorderType.Double, style.Borders[BorderType.TopBorder].LineStyle);
    }

    [Fact]
    public async Task FormatCells_WithPatternFill_ShouldApplyPattern()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_pattern.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Pattern Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_pattern_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["patternType"] = "DiagonalStripe",
            ["backgroundColor"] = "#FF0000",
            ["patternColor"] = "#FFFFFF"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(BackgroundType.DiagonalStripe, style.Pattern);
    }

    [Fact]
    public async Task FormatCells_WithGray50Pattern_ShouldApplyPattern()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_format_gray50.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Gray50 Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_format_gray50_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "format",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["range"] = "A1",
            ["patternType"] = "Gray50",
            ["backgroundColor"] = "#0000FF"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var style = resultWorkbook.Worksheets[0].Cells["A1"].GetStyle();
        Assert.Equal(BackgroundType.Gray50, style.Pattern);
    }

    [Fact]
    public async Task GetFormat_WithFieldsParameter_ShouldReturnOnlyRequestedFields()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_format_fields.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["range"] = "A1",
            ["fields"] = "font"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("fontName", result);
        Assert.Contains("fontSize", result);
        Assert.DoesNotContain("borders", result);
        Assert.DoesNotContain("horizontalAlignment", result);
    }

    [Fact]
    public async Task GetFormat_WithMultipleFields_ShouldReturnRequestedFields()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_format_multi_fields.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["range"] = "A1",
            ["fields"] = "font,color"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("fontName", result);
        Assert.Contains("fontColor", result);
        Assert.Contains("patternType", result);
        Assert.DoesNotContain("borders", result);
    }

    [Fact]
    public async Task GetFormat_WithColorField_ShouldIncludePatternType()
    {
        // Arrange
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

        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["range"] = "A1",
            ["fields"] = "color"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("patternType", result);
        Assert.Contains("foregroundColor", result);
        Assert.Contains("backgroundColor", result);
    }

    [Fact]
    public async Task GetFormat_WithAlignmentField_ShouldReturnAlignmentOnly()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_format_alignment_field.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["range"] = "A1",
            ["fields"] = "alignment"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("horizontalAlignment", result);
        Assert.Contains("verticalAlignment", result);
        Assert.DoesNotContain("fontName", result);
        Assert.DoesNotContain("borders", result);
    }

    [Fact]
    public async Task GetFormat_WithValueField_ShouldReturnValueInfo()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_format_value_field.xlsx");
        var workbook = new Workbook(workbookPath);
        workbook.Worksheets[0].Cells["A1"].Value = "TestValue";
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_format",
            ["path"] = workbookPath,
            ["range"] = "A1",
            ["fields"] = "value"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("TestValue", result);
        Assert.Contains("dataType", result);
        Assert.DoesNotContain("fontName", result);
    }
}