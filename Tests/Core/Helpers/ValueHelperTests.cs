using System.Text.Json.Nodes;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Tests.Core.Helpers;

/// <summary>
///     Unit tests for ValueHelper class
/// </summary>
public class ValueHelperTests
{
    #region ParseValue Tests

    [Fact]
    public void ParseValue_WithInteger_ShouldReturnDouble()
    {
        var result = ValueHelper.ParseValue("42");

        Assert.IsType<double>(result);
        Assert.Equal(42.0, result);
    }

    [Fact]
    public void ParseValue_WithDecimal_ShouldReturnDouble()
    {
        var result = ValueHelper.ParseValue("3.14");

        Assert.IsType<double>(result);
        Assert.Equal(3.14, result);
    }

    [Fact]
    public void ParseValue_WithNegativeNumber_ShouldReturnDouble()
    {
        var result = ValueHelper.ParseValue("-123.45");

        Assert.IsType<double>(result);
        Assert.Equal(-123.45, result);
    }

    [Fact]
    public void ParseValue_WithScientificNotation_ShouldReturnDouble()
    {
        var result = ValueHelper.ParseValue("1.5E+10");

        Assert.IsType<double>(result);
        Assert.Equal(1.5E+10, result);
    }

    [Fact]
    public void ParseValue_WithTrue_ShouldReturnBool()
    {
        var result = ValueHelper.ParseValue("true");

        Assert.IsType<bool>(result);
        Assert.True((bool)result);
    }

    [Fact]
    public void ParseValue_WithFalse_ShouldReturnBool()
    {
        var result = ValueHelper.ParseValue("false");

        Assert.IsType<bool>(result);
        Assert.False((bool)result);
    }

    [Fact]
    public void ParseValue_WithTrueUpperCase_ShouldReturnBool()
    {
        var result = ValueHelper.ParseValue("TRUE");

        Assert.IsType<bool>(result);
        Assert.True((bool)result);
    }

    [Fact]
    public void ParseValue_WithDate_ShouldReturnDateTime()
    {
        var result = ValueHelper.ParseValue("2024-01-15");

        Assert.IsType<DateTime>(result);
        Assert.Equal(new DateTime(2024, 1, 15), result);
    }

    [Fact]
    public void ParseValue_WithDateTime_ShouldReturnDateTime()
    {
        var result = ValueHelper.ParseValue("2024-01-15 10:30:00");

        Assert.IsType<DateTime>(result);
        var dt = (DateTime)result;
        Assert.Equal(2024, dt.Year);
        Assert.Equal(1, dt.Month);
        Assert.Equal(15, dt.Day);
    }

    [Fact]
    public void ParseValue_WithString_ShouldReturnString()
    {
        var result = ValueHelper.ParseValue("Hello World");

        Assert.IsType<string>(result);
        Assert.Equal("Hello World", result);
    }

    [Fact]
    public void ParseValue_WithEmptyString_ShouldReturnEmptyString()
    {
        var result = ValueHelper.ParseValue("");

        Assert.IsType<string>(result);
        Assert.Equal("", result);
    }

    #endregion

    #region GetArray Tests

    [Fact]
    public void GetArray_WithValidArray_ShouldReturnJsonArray()
    {
        var obj = JsonNode.Parse("{\"items\": [1, 2, 3]}")?.AsObject();

        var result = ValueHelper.GetArray(obj, "items");

        Assert.NotNull(result);
        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void GetArray_WithMissingKey_ShouldReturnNull()
    {
        var obj = JsonNode.Parse("{\"other\": \"value\"}")?.AsObject();

        var result = ValueHelper.GetArray(obj, "items");

        Assert.Null(result);
    }

    [Fact]
    public void GetArray_WithNullObject_ShouldReturnNull()
    {
        var result = ValueHelper.GetArray(null, "items");

        Assert.Null(result);
    }

    [Fact]
    public void GetArray_WithNonArrayValue_ShouldReturnNull()
    {
        var obj = JsonNode.Parse("{\"items\": \"not an array\"}")?.AsObject();

        var result = ValueHelper.GetArray(obj, "items");

        Assert.Null(result);
    }

    #endregion

    #region GetString Tests

    [Fact]
    public void GetString_WithValidString_ShouldReturnValue()
    {
        var obj = JsonNode.Parse("{\"name\": \"test\"}")?.AsObject();

        var result = ValueHelper.GetString(obj, "name");

        Assert.Equal("test", result);
    }

    [Fact]
    public void GetString_WithMissingKey_ShouldReturnDefault()
    {
        var obj = JsonNode.Parse("{\"other\": \"value\"}")?.AsObject();

        var result = ValueHelper.GetString(obj, "name");

        Assert.Equal("", result);
    }

    [Fact]
    public void GetString_WithMissingKey_ShouldReturnCustomDefault()
    {
        var obj = JsonNode.Parse("{\"other\": \"value\"}")?.AsObject();

        var result = ValueHelper.GetString(obj, "name", "default_value");

        Assert.Equal("default_value", result);
    }

    [Fact]
    public void GetString_WithNullObject_ShouldReturnDefault()
    {
        var result = ValueHelper.GetString(null, "name");

        Assert.Equal("", result);
    }

    [Fact]
    public void GetString_WithNullObject_ShouldReturnCustomDefault()
    {
        var result = ValueHelper.GetString(null, "name", "custom");

        Assert.Equal("custom", result);
    }

    [Fact]
    public void GetString_WithNullValue_ShouldReturnDefault()
    {
        var obj = JsonNode.Parse("{\"name\": null}")?.AsObject();

        var result = ValueHelper.GetString(obj, "name", "default");

        Assert.Equal("default", result);
    }

    #endregion
}