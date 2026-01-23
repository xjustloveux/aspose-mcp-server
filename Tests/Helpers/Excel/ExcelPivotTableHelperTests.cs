using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.Excel;

public class ExcelPivotTableHelperTests : ExcelTestBase
{
    #region ParsePivotTableStyle Tests

    [Theory]
    [InlineData("None", PivotTableStyleType.None)]
    [InlineData("none", PivotTableStyleType.None)]
    [InlineData("NONE", PivotTableStyleType.None)]
    public void ParsePivotTableStyle_WithNone_ReturnsNone(string input, PivotTableStyleType expected)
    {
        var result = ExcelPivotTableHelper.ParsePivotTableStyle(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("Light1", PivotTableStyleType.PivotTableStyleLight1)]
    [InlineData("Medium1", PivotTableStyleType.PivotTableStyleMedium1)]
    [InlineData("Dark1", PivotTableStyleType.PivotTableStyleDark1)]
    public void ParsePivotTableStyle_WithStyleNames_ReturnsCorrectStyle(string input, PivotTableStyleType expected)
    {
        var result = ExcelPivotTableHelper.ParsePivotTableStyle(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("PivotTableStyleLight1", PivotTableStyleType.PivotTableStyleLight1)]
    [InlineData("PivotTableStyleMedium1", PivotTableStyleType.PivotTableStyleMedium1)]
    [InlineData("PivotTableStyleDark1", PivotTableStyleType.PivotTableStyleDark1)]
    public void ParsePivotTableStyle_WithFullStyleNames_ReturnsCorrectStyle(string input, PivotTableStyleType expected)
    {
        var result = ExcelPivotTableHelper.ParsePivotTableStyle(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    public void ParsePivotTableStyle_WithInvalidStyle_ThrowsArgumentException(string input)
    {
        var ex = Assert.Throws<ArgumentException>(() => ExcelPivotTableHelper.ParsePivotTableStyle(input));

        Assert.Contains("Invalid style", ex.Message);
    }

    #endregion

    #region FindFieldIndex Tests

    [Fact]
    public void FindFieldIndex_WithValidFieldName_ReturnsIndex()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["B1"].Value = "Value";
        worksheet.Cells["A2"].Value = "Item1";
        worksheet.Cells["B2"].Value = "100";
        var range = worksheet.Cells.CreateRange("A1:B2");

        var result = ExcelPivotTableHelper.FindFieldIndex(worksheet, range, "Name");

        Assert.Equal(0, result);
    }

    [Fact]
    public void FindFieldIndex_WithSecondColumn_ReturnsCorrectIndex()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["B1"].Value = "Value";
        var range = worksheet.Cells.CreateRange("A1:B1");

        var result = ExcelPivotTableHelper.FindFieldIndex(worksheet, range, "Value");

        Assert.Equal(1, result);
    }

    [Fact]
    public void FindFieldIndex_WithNonExistentField_ReturnsMinusOne()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Name";
        var range = worksheet.Cells.CreateRange("A1:A1");

        var result = ExcelPivotTableHelper.FindFieldIndex(worksheet, range, "NotFound");

        Assert.Equal(-1, result);
    }

    [Fact]
    public void FindFieldIndex_WithFieldInDataArea_ReturnsIndex()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Header";
        worksheet.Cells["A2"].Value = "SearchThis";
        var range = worksheet.Cells.CreateRange("A1:A2");

        var result = ExcelPivotTableHelper.FindFieldIndex(worksheet, range, "SearchThis");

        Assert.Equal(0, result);
    }

    #endregion

    #region GetAvailableFieldNames Tests

    [Fact]
    public void GetAvailableFieldNames_WithHeaders_ReturnsHeaderNames()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["B1"].Value = "Value";
        worksheet.Cells["C1"].Value = "Date";
        var range = worksheet.Cells.CreateRange("A1:C2");

        var result = ExcelPivotTableHelper.GetAvailableFieldNames(worksheet, range);

        Assert.Equal(3, result.Count);
        Assert.Contains("Name", result);
        Assert.Contains("Value", result);
        Assert.Contains("Date", result);
    }

    [Fact]
    public void GetAvailableFieldNames_WithEmptyHeaders_ReturnsNonEmptyOnly()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].Value = "Name";
        worksheet.Cells["C1"].Value = "Date";
        var range = worksheet.Cells.CreateRange("A1:C1");

        var result = ExcelPivotTableHelper.GetAvailableFieldNames(worksheet, range);

        Assert.Equal(2, result.Count);
        Assert.Contains("Name", result);
        Assert.Contains("Date", result);
    }

    [Fact]
    public void GetAvailableFieldNames_WithNoHeaders_ReturnsEmptyList()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        var range = worksheet.Cells.CreateRange("A1:C1");

        var result = ExcelPivotTableHelper.GetAvailableFieldNames(worksheet, range);

        Assert.Empty(result);
    }

    #endregion

    #region ParseFieldType Tests

    [Theory]
    [InlineData("row", PivotFieldType.Row)]
    [InlineData("Row", PivotFieldType.Row)]
    [InlineData("ROW", PivotFieldType.Row)]
    [InlineData("column", PivotFieldType.Column)]
    [InlineData("Column", PivotFieldType.Column)]
    [InlineData("data", PivotFieldType.Data)]
    [InlineData("Data", PivotFieldType.Data)]
    [InlineData("page", PivotFieldType.Page)]
    [InlineData("Page", PivotFieldType.Page)]
    public void ParseFieldType_WithValidValues_ReturnsCorrectType(string input, PivotFieldType expected)
    {
        var result = ExcelPivotTableHelper.ParseFieldType(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    public void ParseFieldType_WithInvalidValues_ThrowsArgumentException(string input)
    {
        var ex = Assert.Throws<ArgumentException>(() => ExcelPivotTableHelper.ParseFieldType(input));

        Assert.Contains("Invalid fieldType", ex.Message);
    }

    #endregion

    #region ParseFunction Tests

    [Theory]
    [InlineData("Count", ConsolidationFunction.Count)]
    [InlineData("Average", ConsolidationFunction.Average)]
    [InlineData("Max", ConsolidationFunction.Max)]
    [InlineData("Min", ConsolidationFunction.Min)]
    public void ParseFunction_WithValidValues_ReturnsCorrectFunction(string input, ConsolidationFunction expected)
    {
        var result = ExcelPivotTableHelper.ParseFunction(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("Sum")]
    [InlineData("invalid")]
    [InlineData("")]
    public void ParseFunction_WithInvalidOrDefaultValues_ReturnsSum(string input)
    {
        var result = ExcelPivotTableHelper.ParseFunction(input);

        Assert.Equal(ConsolidationFunction.Sum, result);
    }

    #endregion
}
