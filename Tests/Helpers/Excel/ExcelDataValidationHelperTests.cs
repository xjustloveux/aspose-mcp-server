using Aspose.Cells;
using AsposeMcpServer.Helpers.Excel;

namespace AsposeMcpServer.Tests.Helpers.Excel;

public class ExcelDataValidationHelperTests
{
    #region ParseValidationType Tests

    [Theory]
    [InlineData("WholeNumber", ValidationType.WholeNumber)]
    [InlineData("Decimal", ValidationType.Decimal)]
    [InlineData("List", ValidationType.List)]
    [InlineData("Date", ValidationType.Date)]
    [InlineData("Time", ValidationType.Time)]
    [InlineData("TextLength", ValidationType.TextLength)]
    [InlineData("Custom", ValidationType.Custom)]
    public void ParseValidationType_WithValidValues_ReturnsCorrectType(string input, ValidationType expected)
    {
        var result = ExcelDataValidationHelper.ParseValidationType(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("Invalid")]
    [InlineData("wholenumber")]
    [InlineData("")]
    [InlineData("Unknown")]
    public void ParseValidationType_WithInvalidValues_ThrowsArgumentException(string input)
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelDataValidationHelper.ParseValidationType(input));

        Assert.Contains("Unsupported validation type", ex.Message);
    }

    #endregion

    #region ParseOperatorType Tests

    [Theory]
    [InlineData("Between", OperatorType.Between)]
    [InlineData("Equal", OperatorType.Equal)]
    [InlineData("NotEqual", OperatorType.NotEqual)]
    [InlineData("GreaterThan", OperatorType.GreaterThan)]
    [InlineData("LessThan", OperatorType.LessThan)]
    [InlineData("GreaterOrEqual", OperatorType.GreaterOrEqual)]
    [InlineData("LessOrEqual", OperatorType.LessOrEqual)]
    public void ParseOperatorType_WithValidValues_ReturnsCorrectType(string input, OperatorType expected)
    {
        var result = ExcelDataValidationHelper.ParseOperatorType(input, null);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ParseOperatorType_WithNullOperatorAndFormula2_ReturnsBetween()
    {
        var result = ExcelDataValidationHelper.ParseOperatorType(null, "10");

        Assert.Equal(OperatorType.Between, result);
    }

    [Fact]
    public void ParseOperatorType_WithEmptyOperatorAndFormula2_ReturnsBetween()
    {
        var result = ExcelDataValidationHelper.ParseOperatorType("", "20");

        Assert.Equal(OperatorType.Between, result);
    }

    [Fact]
    public void ParseOperatorType_WithNullOperatorAndNoFormula2_ReturnsEqual()
    {
        var result = ExcelDataValidationHelper.ParseOperatorType(null, null);

        Assert.Equal(OperatorType.Equal, result);
    }

    [Fact]
    public void ParseOperatorType_WithEmptyOperatorAndEmptyFormula2_ReturnsEqual()
    {
        var result = ExcelDataValidationHelper.ParseOperatorType("", "");

        Assert.Equal(OperatorType.Equal, result);
    }

    [Theory]
    [InlineData("Invalid")]
    [InlineData("between")]
    [InlineData("equals")]
    public void ParseOperatorType_WithInvalidOperator_ThrowsArgumentException(string input)
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelDataValidationHelper.ParseOperatorType(input, null));

        Assert.Contains("Unsupported operator type", ex.Message);
    }

    #endregion

    #region ValidateCollectionIndex Tests

    [Theory]
    [InlineData(0, 5)]
    [InlineData(1, 5)]
    [InlineData(4, 5)]
    public void ValidateCollectionIndex_WithValidIndex_DoesNotThrow(int index, int count)
    {
        var exception = Record.Exception(() =>
            ExcelDataValidationHelper.ValidateCollectionIndex(index, count, "validation"));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateCollectionIndex_WithNegativeIndex_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelDataValidationHelper.ValidateCollectionIndex(-1, 5, "validation"));

        Assert.Contains("index -1 is out of range", ex.Message);
        Assert.Contains("validation", ex.Message);
    }

    [Fact]
    public void ValidateCollectionIndex_WithIndexEqualToCount_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelDataValidationHelper.ValidateCollectionIndex(5, 5, "validation"));

        Assert.Contains("index 5 is out of range", ex.Message);
    }

    [Fact]
    public void ValidateCollectionIndex_WithIndexGreaterThanCount_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelDataValidationHelper.ValidateCollectionIndex(10, 3, "chart"));

        Assert.Contains("index 10 is out of range", ex.Message);
        Assert.Contains("collection has 3 charts", ex.Message);
    }

    [Fact]
    public void ValidateCollectionIndex_WithEmptyCollection_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelDataValidationHelper.ValidateCollectionIndex(0, 0, "item"));

        Assert.Contains("index 0 is out of range", ex.Message);
        Assert.Contains("collection has 0 items", ex.Message);
    }

    #endregion
}
