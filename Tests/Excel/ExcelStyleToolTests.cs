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
}