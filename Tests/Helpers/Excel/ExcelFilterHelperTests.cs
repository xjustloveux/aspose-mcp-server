using Aspose.Cells;
using AsposeMcpServer.Helpers.Excel;

namespace AsposeMcpServer.Tests.Helpers.Excel;

public class ExcelFilterHelperTests
{
    #region IsNumericOperator Tests

    [Theory]
    [InlineData(FilterOperatorType.GreaterThan, true)]
    [InlineData(FilterOperatorType.GreaterOrEqual, true)]
    [InlineData(FilterOperatorType.LessThan, true)]
    [InlineData(FilterOperatorType.LessOrEqual, true)]
    public void IsNumericOperator_WithNumericOperators_ReturnsTrue(FilterOperatorType op, bool expected)
    {
        var result = ExcelFilterHelper.IsNumericOperator(op);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(FilterOperatorType.Equal)]
    [InlineData(FilterOperatorType.NotEqual)]
    [InlineData(FilterOperatorType.Contains)]
    [InlineData(FilterOperatorType.NotContains)]
    [InlineData(FilterOperatorType.BeginsWith)]
    [InlineData(FilterOperatorType.EndsWith)]
    public void IsNumericOperator_WithNonNumericOperators_ReturnsFalse(FilterOperatorType op)
    {
        var result = ExcelFilterHelper.IsNumericOperator(op);

        Assert.False(result);
    }

    #endregion

    #region ParseFilterOperator Tests

    [Theory]
    [InlineData("Equal", FilterOperatorType.Equal)]
    [InlineData("NotEqual", FilterOperatorType.NotEqual)]
    [InlineData("GreaterThan", FilterOperatorType.GreaterThan)]
    [InlineData("GreaterOrEqual", FilterOperatorType.GreaterOrEqual)]
    [InlineData("LessThan", FilterOperatorType.LessThan)]
    [InlineData("LessOrEqual", FilterOperatorType.LessOrEqual)]
    [InlineData("Contains", FilterOperatorType.Contains)]
    [InlineData("NotContains", FilterOperatorType.NotContains)]
    [InlineData("BeginsWith", FilterOperatorType.BeginsWith)]
    [InlineData("EndsWith", FilterOperatorType.EndsWith)]
    public void ParseFilterOperator_WithValidValues_ReturnsCorrectOperator(string input, FilterOperatorType expected)
    {
        var result = ExcelFilterHelper.ParseFilterOperator(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("Invalid")]
    [InlineData("equal")]
    [InlineData("")]
    [InlineData("Equals")]
    [InlineData("LikesWith")]
    public void ParseFilterOperator_WithInvalidValues_ThrowsArgumentException(string input)
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ExcelFilterHelper.ParseFilterOperator(input));

        Assert.Contains("Unsupported filter operator", ex.Message);
    }

    #endregion
}
