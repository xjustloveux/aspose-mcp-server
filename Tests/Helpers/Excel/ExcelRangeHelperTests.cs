using Aspose.Cells;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.Excel;

public class ExcelRangeHelperTests : ExcelTestBase
{
    #region TextFormatNumber Constant

    [Fact]
    public void TextFormatNumber_IsCorrectValue()
    {
        Assert.Equal(49, ExcelRangeHelper.TextFormatNumber);
    }

    #endregion

    #region ParseDataArray Tests

    [Fact]
    public void ParseDataArray_WithValidArray_ReturnsJsonArray()
    {
        var json = "[[1, 2], [3, 4]]";

        var result = ExcelRangeHelper.ParseDataArray(json);

        Assert.NotNull(result);
        Assert.Equal(2, result.Count);
        Assert.Equal(2, result[0]!.AsArray().Count);
        Assert.Equal(2, result[1]!.AsArray().Count);
    }

    [Fact]
    public void ParseDataArray_WithEmptyArray_ReturnsEmptyJsonArray()
    {
        var json = "[]";

        var result = ExcelRangeHelper.ParseDataArray(json);

        Assert.NotNull(result);
        Assert.Empty(result);
    }

    [Fact]
    public void ParseDataArray_WithInvalidJson_ThrowsArgumentException()
    {
        var json = "not valid json";

        var ex = Assert.Throws<ArgumentException>(() => ExcelRangeHelper.ParseDataArray(json));

        Assert.Contains("Invalid JSON format", ex.Message);
    }

    [Fact]
    public void ParseDataArray_WithNonArrayJson_ThrowsInvalidOperationException()
    {
        var json = "{\"key\": \"value\"}";

        Assert.Throws<InvalidOperationException>(() => ExcelRangeHelper.ParseDataArray(json));
    }

    #endregion

    #region LooksLikeCellReference Tests

    [Theory]
    [InlineData("A1", true)]
    [InlineData("B2", true)]
    [InlineData("Z9", true)]
    [InlineData("AA1", true)]
    [InlineData("AB12", true)]
    public void LooksLikeCellReference_WithValidCellRefs_ReturnsTrue(string input, bool expected)
    {
        var result = ExcelRangeHelper.LooksLikeCellReference(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("A")]
    [InlineData("1")]
    [InlineData("Hello")]
    [InlineData("A:B")]
    [InlineData("$A$1")]
    [InlineData("A 1")]
    [InlineData("")]
    public void LooksLikeCellReference_WithInvalidCellRefs_ReturnsFalse(string input)
    {
        var result = ExcelRangeHelper.LooksLikeCellReference(input);

        Assert.False(result);
    }

    [Theory]
    [InlineData("ABCDEF1")]
    [InlineData("A123456")]
    public void LooksLikeCellReference_WithTooLongRefs_ReturnsFalse(string input)
    {
        var result = ExcelRangeHelper.LooksLikeCellReference(input);

        Assert.False(result);
    }

    #endregion

    #region Write2DArrayData Tests

    [Fact]
    public void Write2DArrayData_WithValidData_WritesToWorksheet()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var dataArray = ExcelRangeHelper.ParseDataArray("[[\"A\", \"B\"], [\"C\", \"D\"]]");

        ExcelRangeHelper.Write2DArrayData(workbook, worksheet, 0, 0, dataArray);

        Assert.Equal("A", worksheet.Cells["A1"].Value?.ToString());
        Assert.Equal("B", worksheet.Cells["B1"].Value?.ToString());
        Assert.Equal("C", worksheet.Cells["A2"].Value?.ToString());
        Assert.Equal("D", worksheet.Cells["B2"].Value?.ToString());
    }

    [Fact]
    public void Write2DArrayData_WithEmptyArray_ThrowsException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var dataArray = ExcelRangeHelper.ParseDataArray("[]");

        Assert.Throws<InvalidOperationException>(() =>
            ExcelRangeHelper.Write2DArrayData(workbook, worksheet, 0, 0, dataArray));
    }

    [Fact]
    public void Write2DArrayData_WithStartOffset_WritesAtCorrectPosition()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var dataArray = ExcelRangeHelper.ParseDataArray("[[\"Test\"]]");

        ExcelRangeHelper.Write2DArrayData(workbook, worksheet, 2, 3, dataArray);

        Assert.Null(worksheet.Cells["A1"].Value);
        Assert.Equal("Test", worksheet.Cells[2, 3].Value?.ToString());
    }

    #endregion

    #region WriteObjectArrayData Tests

    [Fact]
    public void WriteObjectArrayData_WithValidData_WritesToCells()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var dataArray = ExcelRangeHelper.ParseDataArray(
            "[{\"cell\": \"A1\", \"value\": \"Hello\"}, {\"cell\": \"B2\", \"value\": \"World\"}]");

        ExcelRangeHelper.WriteObjectArrayData(worksheet, dataArray);

        Assert.Equal("Hello", worksheet.Cells["A1"].Value?.ToString());
        Assert.Equal("World", worksheet.Cells["B2"].Value?.ToString());
    }

    [Fact]
    public void WriteObjectArrayData_WithEmptyValue_WritesEmptyString()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var dataArray = ExcelRangeHelper.ParseDataArray("[{\"cell\": \"A1\", \"value\": \"\"}]");

        ExcelRangeHelper.WriteObjectArrayData(worksheet, dataArray);

        Assert.Equal(string.Empty, worksheet.Cells["A1"].Value?.ToString() ?? string.Empty);
    }

    [Fact]
    public void WriteObjectArrayData_WithInvalidFormat_ThrowsArgumentException()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var dataArray = ExcelRangeHelper.ParseDataArray("[\"invalid\"]");

        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelRangeHelper.WriteObjectArrayData(worksheet, dataArray));

        Assert.Contains("Invalid data format", ex.Message);
    }

    #endregion

    #region GetPasteType Tests

    [Theory]
    [InlineData("Values", PasteType.Values)]
    [InlineData("Formats", PasteType.Formats)]
    [InlineData("Formulas", PasteType.Formulas)]
    public void GetPasteType_WithValidValues_ReturnsCorrectType(string input, PasteType expected)
    {
        var result = ExcelRangeHelper.GetPasteType(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("All")]
    [InlineData("invalid")]
    [InlineData("")]
    public void GetPasteType_WithInvalidOrDefaultValues_ReturnsAll(string input)
    {
        var result = ExcelRangeHelper.GetPasteType(input);

        Assert.Equal(PasteType.All, result);
    }

    #endregion
}
