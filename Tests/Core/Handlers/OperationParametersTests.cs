using System.Text.Json;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Tests.Core.Handlers;

/// <summary>
///     Unit tests for OperationParameters
/// </summary>
public class OperationParametersTests
{
    public enum TestEnum
    {
        Value1,
        Value2,
        Value3
    }

    [Fact]
    public void Set_And_Has_WorksCorrectly()
    {
        var parameters = new OperationParameters();

        parameters.Set("name", "test");

        Assert.True(parameters.Has("name"));
        Assert.False(parameters.Has("other"));
    }

    [Fact]
    public void Has_WithNullValue_ReturnsFalse()
    {
        var parameters = new OperationParameters();

        parameters.Set("name", null);

        Assert.False(parameters.Has("name"));
    }

    [Fact]
    public void Has_IsCaseInsensitive()
    {
        var parameters = new OperationParameters();

        parameters.Set("TestName", "value");

        Assert.True(parameters.Has("testname"));
        Assert.True(parameters.Has("TESTNAME"));
        Assert.True(parameters.Has("TestName"));
    }

    [Fact]
    public void GetRequired_WithValue_ReturnsValue()
    {
        var parameters = new OperationParameters();
        parameters.Set("text", "hello");

        var result = parameters.GetRequired<string>("text");

        Assert.Equal("hello", result);
    }

    [Fact]
    public void GetRequired_WithMissingValue_ThrowsArgumentException()
    {
        var parameters = new OperationParameters();

        var ex = Assert.Throws<ArgumentException>(() => parameters.GetRequired<string>("missing"));

        Assert.Contains("missing", ex.Message);
        Assert.Contains("is required", ex.Message);
    }

    [Fact]
    public void GetRequired_WithNullValue_ThrowsArgumentException()
    {
        var parameters = new OperationParameters();
        parameters.Set("name", null);

        var ex = Assert.Throws<ArgumentException>(() => parameters.GetRequired<string>("name"));

        Assert.Contains("name", ex.Message);
    }

    [Fact]
    public void GetRequired_IsCaseInsensitive()
    {
        var parameters = new OperationParameters();
        parameters.Set("MyParam", "value");

        Assert.Equal("value", parameters.GetRequired<string>("myparam"));
        Assert.Equal("value", parameters.GetRequired<string>("MYPARAM"));
    }

    [Fact]
    public void GetOptional_WithValue_ReturnsValue()
    {
        var parameters = new OperationParameters();
        parameters.Set("count", 42);

        var result = parameters.GetOptional("count", 0);

        Assert.Equal(42, result);
    }

    [Fact]
    public void GetOptional_WithMissingValue_ReturnsDefault()
    {
        var parameters = new OperationParameters();

        var result = parameters.GetOptional("missing", 100);

        Assert.Equal(100, result);
    }

    [Fact]
    public void GetOptional_WithNullValue_ReturnsDefault()
    {
        var parameters = new OperationParameters();
        parameters.Set("name", null);

        var result = parameters.GetOptional("name", "default");

        Assert.Equal("default", result);
    }

    [Fact]
    public void GetOptional_NullableInt_WithValue_ReturnsValue()
    {
        var parameters = new OperationParameters();
        parameters.Set("count", 42);

        var result = parameters.GetOptional<int?>("count");

        Assert.Equal(42, result);
    }

    [Fact]
    public void GetOptional_NullableInt_WithMissingValue_ReturnsNull()
    {
        var parameters = new OperationParameters();

        var result = parameters.GetOptional<int?>("missing");

        Assert.Null(result);
    }

    [Fact]
    public void GetOptional_NullableDouble_WithValue_ReturnsValue()
    {
        var parameters = new OperationParameters();
        parameters.Set("size", 12.5);

        var result = parameters.GetOptional<double?>("size");

        Assert.Equal(12.5, result);
    }

    [Fact]
    public void GetOptional_NullableBool_WithValue_ReturnsValue()
    {
        var parameters = new OperationParameters();
        parameters.Set("enabled", true);

        var result = parameters.GetOptional<bool?>("enabled");

        Assert.True(result);
    }

    [Fact]
    public void GetRequired_ConvertsIntToDouble()
    {
        var parameters = new OperationParameters();
        parameters.Set("value", 42);

        var result = parameters.GetRequired<double>("value");

        Assert.Equal(42.0, result);
    }

    [Fact]
    public void GetRequired_ConvertsStringToInt()
    {
        var parameters = new OperationParameters();
        parameters.Set("value", "123");

        var result = parameters.GetRequired<int>("value");

        Assert.Equal(123, result);
    }

    [Fact]
    public void GetRequired_Enum_FromString()
    {
        var parameters = new OperationParameters();
        parameters.Set("type", "Value2");

        var result = parameters.GetRequired<TestEnum>("type");

        Assert.Equal(TestEnum.Value2, result);
    }

    [Fact]
    public void GetRequired_Enum_FromStringCaseInsensitive()
    {
        var parameters = new OperationParameters();
        parameters.Set("type", "value3");

        var result = parameters.GetRequired<TestEnum>("type");

        Assert.Equal(TestEnum.Value3, result);
    }

    [Fact]
    public void GetRequired_Enum_FromInt()
    {
        var parameters = new OperationParameters();
        parameters.Set("type", 1);

        var result = parameters.GetRequired<TestEnum>("type");

        Assert.Equal(TestEnum.Value2, result);
    }

    [Fact]
    public void GetOptional_NullableEnum_WithValue_ReturnsValue()
    {
        var parameters = new OperationParameters();
        parameters.Set("type", "Value1");

        var result = parameters.GetOptional<TestEnum?>("type");

        Assert.Equal(TestEnum.Value1, result);
    }

    [Fact]
    public void GetOptional_NullableEnum_WithMissingValue_ReturnsNull()
    {
        var parameters = new OperationParameters();

        var result = parameters.GetOptional<TestEnum?>("type");

        Assert.Null(result);
    }

    [Fact]
    public void GetRequired_JsonElement_String()
    {
        var parameters = new OperationParameters();
        var json = JsonDocument.Parse("\"hello\"");
        parameters.Set("text", json.RootElement);

        var result = parameters.GetRequired<string>("text");

        Assert.Equal("hello", result);
    }

    [Fact]
    public void GetRequired_JsonElement_Int()
    {
        var parameters = new OperationParameters();
        var json = JsonDocument.Parse("42");
        parameters.Set("count", json.RootElement);

        var result = parameters.GetRequired<int>("count");

        Assert.Equal(42, result);
    }

    [Fact]
    public void GetRequired_JsonElement_Double()
    {
        var parameters = new OperationParameters();
        var json = JsonDocument.Parse("3.14");
        parameters.Set("value", json.RootElement);

        var result = parameters.GetRequired<double>("value");

        Assert.Equal(3.14, result);
    }

    [Fact]
    public void GetRequired_JsonElement_Bool()
    {
        var parameters = new OperationParameters();
        var json = JsonDocument.Parse("true");
        parameters.Set("enabled", json.RootElement);

        var result = parameters.GetRequired<bool>("enabled");

        Assert.True(result);
    }

    [Fact]
    public void GetOptional_JsonElement_NullableInt()
    {
        var parameters = new OperationParameters();
        var json = JsonDocument.Parse("99");
        parameters.Set("count", json.RootElement);

        var result = parameters.GetOptional<int?>("count");

        Assert.Equal(99, result);
    }

    [Fact]
    public void GetRequired_JsonElement_Enum_FromString()
    {
        var parameters = new OperationParameters();
        var json = JsonDocument.Parse("\"Value2\"");
        parameters.Set("type", json.RootElement);

        var result = parameters.GetRequired<TestEnum>("type");

        Assert.Equal(TestEnum.Value2, result);
    }

    [Fact]
    public void GetRequired_JsonElement_Enum_FromInt()
    {
        var parameters = new OperationParameters();
        var json = JsonDocument.Parse("2");
        parameters.Set("type", json.RootElement);

        var result = parameters.GetRequired<TestEnum>("type");

        Assert.Equal(TestEnum.Value3, result);
    }

    [Fact]
    public void GetRaw_ReturnsUnconvertedValue()
    {
        var parameters = new OperationParameters();
        var originalValue = new object();
        parameters.Set("obj", originalValue);

        var result = parameters.GetRaw("obj");

        Assert.Same(originalValue, result);
    }

    [Fact]
    public void GetRaw_WithMissingKey_ReturnsNull()
    {
        var parameters = new OperationParameters();

        var result = parameters.GetRaw("missing");

        Assert.Null(result);
    }

    [Fact]
    public void GetRequired_InvalidConversion_ThrowsArgumentException()
    {
        var parameters = new OperationParameters();
        parameters.Set("value", "not-a-number");

        var ex = Assert.Throws<ArgumentException>(() => parameters.GetRequired<int>("value"));

        Assert.Contains("value", ex.Message);
        Assert.Contains("Cannot convert", ex.Message);
    }

    [Fact]
    public void GetRequired_InvalidEnumString_ThrowsArgumentException()
    {
        var parameters = new OperationParameters();
        parameters.Set("type", "InvalidValue");

        var ex = Assert.Throws<ArgumentException>(() => parameters.GetRequired<TestEnum>("type"));

        Assert.Contains("type", ex.Message);
        Assert.Contains("Cannot convert", ex.Message);
    }

    [Fact]
    public void Set_OverwritesExistingValue()
    {
        var parameters = new OperationParameters();
        parameters.Set("key", "original");
        parameters.Set("key", "updated");

        Assert.Equal("updated", parameters.GetRequired<string>("key"));
    }

    [Fact]
    public void GetOptional_WithExplicitNullDefault_ReturnsNull()
    {
        var parameters = new OperationParameters();

        var result = parameters.GetOptional<string?>("missing");

        Assert.Null(result);
    }
}
