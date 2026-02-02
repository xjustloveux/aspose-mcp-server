using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using AsposeMcpServer.Core;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Shape;
using AsposeMcpServer.Results.Word.Field;

namespace AsposeMcpServer.Tests.Core;

/// <summary>
///     Unit tests for OutputSchemaGenerator.
/// </summary>
public class OutputSchemaGeneratorTests
{
    /// <summary>
    ///     Verifies that GenerateForType returns a schema with data and output fields.
    /// </summary>
    [Fact]
    public void GenerateForType_ShouldReturnSchemaWithDataAndOutputFields()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(SuccessResult));

        var json = schema.GetRawText();
        Assert.False(string.IsNullOrEmpty(json));

        Assert.Contains("\"data\"", json);
        Assert.Contains("\"output\"", json);
        Assert.Contains("\"required\"", json);
    }

    /// <summary>
    ///     Verifies that the generated schema has correct structure.
    /// </summary>
    [Fact]
    public void GenerateForType_ShouldHaveCorrectSchemaStructure()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(SuccessResult));

        var schemaNode = JsonNode.Parse(schema.GetRawText())!.AsObject();

        Assert.Equal("object", schemaNode["type"]?.GetValue<string>());
        Assert.NotNull(schemaNode["properties"]);

        var properties = schemaNode["properties"]!.AsObject();
        Assert.True(properties.ContainsKey("data"));
        Assert.True(properties.ContainsKey("output"));

        var required = schemaNode["required"]!.AsArray();
        Assert.Contains("data", required.Select(r => r?.GetValue<string>()));
        Assert.Contains("output", required.Select(r => r?.GetValue<string>()));
    }

    /// <summary>
    ///     Verifies that the output field contains path, sessionId, and isSession.
    /// </summary>
    [Fact]
    public void GenerateForType_OutputFieldShouldContainExpectedProperties()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(SuccessResult));

        var schemaNode = JsonNode.Parse(schema.GetRawText())!.AsObject();
        var outputProperties = schemaNode["properties"]!["output"]!["properties"]!.AsObject();

        Assert.True(outputProperties.ContainsKey("path"));
        Assert.True(outputProperties.ContainsKey("sessionId"));
        Assert.True(outputProperties.ContainsKey("isSession"));
    }

    /// <summary>
    ///     Verifies that the data field contains the result type properties.
    /// </summary>
    [Fact]
    public void GenerateForType_DataFieldShouldContainResultTypeProperties()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(SuccessResult));

        var schemaNode = JsonNode.Parse(schema.GetRawText())!.AsObject();
        var dataProperties = schemaNode["properties"]!["data"]!["properties"]!.AsObject();

        Assert.True(dataProperties.ContainsKey("message"));
    }

    /// <summary>
    ///     Verifies that GenerateForTypes with multiple types returns oneOf schema.
    /// </summary>
    [Fact]
    public void GenerateForTypes_WithMultipleTypes_ShouldReturnOneOfSchema()
    {
        var types = new[] { typeof(TestResultA), typeof(TestResultB) };

        var schema = OutputSchemaGenerator.GenerateForTypes(types);

        var schemaNode = JsonNode.Parse(schema.GetRawText())!.AsObject();
        var dataSchema = schemaNode["properties"]!["data"]!.AsObject();

        Assert.True(dataSchema.ContainsKey("oneOf"));
        var oneOfArray = dataSchema["oneOf"]!.AsArray();
        Assert.Equal(2, oneOfArray.Count);
    }

    /// <summary>
    ///     Verifies that GenerateForTypes with single type does not use oneOf.
    /// </summary>
    [Fact]
    public void GenerateForTypes_WithSingleType_ShouldNotUseOneOf()
    {
        var types = new[] { typeof(SuccessResult) };

        var schema = OutputSchemaGenerator.GenerateForTypes(types);

        var schemaNode = JsonNode.Parse(schema.GetRawText())!.AsObject();
        var dataSchema = schemaNode["properties"]!["data"]!.AsObject();

        Assert.False(dataSchema.ContainsKey("oneOf"));
        Assert.True(dataSchema.ContainsKey("properties"));
    }

    /// <summary>
    ///     Verifies that GenerateForTypes with empty collection throws.
    /// </summary>
    [Fact]
    public void GenerateForTypes_WithEmptyCollection_ShouldThrow()
    {
        Assert.Throws<ArgumentException>(() => OutputSchemaGenerator.GenerateForTypes(Array.Empty<Type>()));
    }

    /// <summary>
    ///     Verifies that GenerateFromNamespace returns null for non-existent namespace.
    /// </summary>
    [Fact]
    public void GenerateFromNamespace_WithNonExistentNamespace_ShouldReturnNull()
    {
        var schema = OutputSchemaGenerator.GenerateFromNamespace("NonExistent.Namespace");

        Assert.Null(schema);
    }

    /// <summary>
    ///     Verifies that GenerateFromNamespace returns schema for valid handler namespace.
    /// </summary>
    [Fact]
    public void GenerateFromNamespace_WithValidNamespace_ShouldReturnSchema()
    {
        var schema = OutputSchemaGenerator.GenerateFromNamespace(
            "AsposeMcpServer.Handlers.Word.Text");

        Assert.NotNull(schema);

        var json = schema.Value.GetRawText();
        Assert.Contains("\"data\"", json);
        Assert.Contains("\"output\"", json);
    }

    /// <summary>
    ///     Verifies that GenerateForType handles types with AllTypes static field.
    /// </summary>
    [Fact]
    public void GenerateForType_WithAllTypesField_ShouldGenerateOneOfSchema()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(TestUnionResult));

        var schemaNode = JsonNode.Parse(schema.GetRawText())!.AsObject();
        var dataSchema = schemaNode["properties"]!["data"]!.AsObject();

        Assert.True(dataSchema.ContainsKey("oneOf"));
        var oneOfArray = dataSchema["oneOf"]!.AsArray();
        Assert.Equal(2, oneOfArray.Count);
    }

    #region Polymorphic Property Schema

    /// <summary>
    ///     Verifies that schema generation includes discriminator values for polymorphic properties.
    /// </summary>
    [Fact]
    public void GenerateForType_WithPolymorphicProperty_ShouldIncludeDiscriminators()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(TestPolymorphicResult));

        var json = schema.GetRawText();

        Assert.Contains("typeA", json);
        Assert.Contains("typeB", json);
    }

    /// <summary>
    ///     Verifies that polymorphic property schema contains anyOf or oneOf structure.
    /// </summary>
    [Fact]
    public void GenerateForType_WithPolymorphicProperty_ShouldContainAnyOfOrOneOf()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(TestPolymorphicResult));

        var json = schema.GetRawText();

        Assert.True(
            json.Contains("anyOf") || json.Contains("oneOf"),
            "Schema should contain anyOf or oneOf for polymorphic property");
    }

    /// <summary>
    ///     Verifies that polymorphic property schema includes properties from derived types.
    /// </summary>
    [Fact]
    public void GenerateForType_WithPolymorphicProperty_ShouldIncludeDerivedTypeProperties()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(TestPolymorphicResult));

        var json = schema.GetRawText();

        Assert.Contains("propA", json);
        Assert.Contains("propB", json);
    }

    #endregion

    #region GetShapeDetailsResult Schema

    /// <summary>
    ///     Verifies that GetShapeDetailsResult schema contains the details field.
    /// </summary>
    [Fact]
    public void GenerateForType_GetShapeDetailsResult_ShouldContainDetailsField()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetShapeDetailsResult));

        var json = schema.GetRawText();

        Assert.Contains("\"details\"", json);
    }

    /// <summary>
    ///     Verifies that GetShapeDetailsResult schema includes all 9 shape detail type discriminators.
    /// </summary>
    /// <param name="discriminator">The type discriminator value to check.</param>
    [Theory]
    [InlineData("autoShape")]
    [InlineData("audio")]
    [InlineData("chart")]
    [InlineData("connector")]
    [InlineData("group")]
    [InlineData("picture")]
    [InlineData("smartArt")]
    [InlineData("table")]
    [InlineData("video")]
    public void GenerateForType_GetShapeDetailsResult_ShouldIncludeShapeDetailDiscriminator(
        string discriminator)
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetShapeDetailsResult));

        var json = schema.GetRawText();

        Assert.Contains(discriminator, json);
    }

    /// <summary>
    ///     Verifies that GetShapeDetailsResult schema includes base shape properties from GetShapeInfo.
    /// </summary>
    /// <param name="propertyName">The base property name to check.</param>
    [Theory]
    [InlineData("index")]
    [InlineData("type")]
    [InlineData("x")]
    [InlineData("y")]
    [InlineData("width")]
    [InlineData("height")]
    [InlineData("rotation")]
    [InlineData("hidden")]
    public void GenerateForType_GetShapeDetailsResult_ShouldIncludeBaseProperty(string propertyName)
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetShapeDetailsResult));

        var json = schema.GetRawText();

        Assert.Contains($"\"{propertyName}\"", json);
    }

    /// <summary>
    ///     Verifies that GetShapeDetailsResult schema includes key properties from AutoShapeDetails.
    /// </summary>
    [Fact]
    public void GenerateForType_GetShapeDetailsResult_ShouldIncludeAutoShapeProperties()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetShapeDetailsResult));

        var json = schema.GetRawText();

        Assert.Contains("shapeType", json);
        Assert.Contains("hasTextFrame", json);
        Assert.Contains("fillType", json);
    }

    /// <summary>
    ///     Verifies that GetShapeDetailsResult schema includes key properties from TableDetails.
    /// </summary>
    [Fact]
    public void GenerateForType_GetShapeDetailsResult_ShouldIncludeTableProperties()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetShapeDetailsResult));

        var json = schema.GetRawText();

        Assert.Contains("rows", json);
        Assert.Contains("columns", json);
        Assert.Contains("mergedCellCount", json);
    }

    /// <summary>
    ///     Verifies that GetShapeDetailsResult schema includes key properties from ChartDetails.
    /// </summary>
    [Fact]
    public void GenerateForType_GetShapeDetailsResult_ShouldIncludeChartProperties()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetShapeDetailsResult));

        var json = schema.GetRawText();

        Assert.Contains("chartType", json);
        Assert.Contains("seriesCount", json);
        Assert.Contains("hasDataTable", json);
    }

    #endregion

    #region GetFormFieldsWordResult Schema

    /// <summary>
    ///     Verifies that GetFormFieldsWordResult schema contains the formFields field.
    /// </summary>
    [Fact]
    public void GenerateForType_GetFormFieldsWordResult_ShouldContainFormFieldsField()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetFormFieldsWordResult));

        var json = schema.GetRawText();

        Assert.Contains("\"formFields\"", json);
    }

    /// <summary>
    ///     Verifies that GetFormFieldsWordResult schema includes all 3 form field type discriminators.
    /// </summary>
    /// <param name="discriminator">The type discriminator value to check.</param>
    [Theory]
    [InlineData("text")]
    [InlineData("checkbox")]
    [InlineData("dropdown")]
    public void GenerateForType_GetFormFieldsWordResult_ShouldIncludeFormFieldDiscriminator(
        string discriminator)
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetFormFieldsWordResult));

        var json = schema.GetRawText();

        Assert.Contains(discriminator, json);
    }

    /// <summary>
    ///     Verifies that GetFormFieldsWordResult schema includes base form field properties.
    /// </summary>
    /// <param name="propertyName">The base property name to check.</param>
    [Theory]
    [InlineData("index")]
    [InlineData("name")]
    [InlineData("type")]
    public void GenerateForType_GetFormFieldsWordResult_ShouldIncludeBaseProperty(string propertyName)
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetFormFieldsWordResult));

        var json = schema.GetRawText();

        Assert.Contains($"\"{propertyName}\"", json);
    }

    /// <summary>
    ///     Verifies that GetFormFieldsWordResult schema includes derived type specific properties.
    /// </summary>
    [Fact]
    public void GenerateForType_GetFormFieldsWordResult_ShouldIncludeDerivedTypeProperties()
    {
        var schema = OutputSchemaGenerator.GenerateForType(typeof(GetFormFieldsWordResult));

        var json = schema.GetRawText();

        Assert.Contains("value", json);
        Assert.Contains("isChecked", json);
        Assert.Contains("selectedIndex", json);
        Assert.Contains("options", json);
    }

    #endregion

    /// <summary>
    ///     Test result type A for JSON schema generation testing.
    /// </summary>
    // ReSharper disable UnusedMember.Local - Type and property used via reflection for JSON schema generation
    private record TestResultA
    {
        [JsonPropertyName("valueA")] public string? ValueA { get; init; }
    }
    // ReSharper restore UnusedMember.Local

    /// <summary>
    ///     Test result type B for JSON schema generation testing.
    /// </summary>
    // ReSharper disable UnusedMember.Local - Type and property used via reflection for JSON schema generation
    private record TestResultB
    {
        [JsonPropertyName("valueB")] public int ValueB { get; init; }
    }
    // ReSharper restore UnusedMember.Local

    /// <summary>
    ///     Test union result type with AllTypes field for reflection-based schema generation.
    /// </summary>
    // ReSharper disable UnusedMember.Local - Type and field accessed via reflection
    private class TestUnionResult
    {
        public static readonly Type[] AllTypes = [typeof(TestResultA), typeof(TestResultB)];
    }
    // ReSharper restore UnusedMember.Local

    /// <summary>
    ///     Abstract base type for testing polymorphic schema generation.
    /// </summary>
    [JsonPolymorphic(TypeDiscriminatorPropertyName = "$type")]
    [JsonDerivedType(typeof(TestDerivedA), "typeA")]
    [JsonDerivedType(typeof(TestDerivedB), "typeB")]
    // ReSharper disable UnusedMember.Local - Types used via reflection for JSON schema generation
    private abstract record TestPolymorphicBase;

    /// <summary>
    ///     Test derived type A for polymorphic schema testing.
    /// </summary>
    private sealed record TestDerivedA : TestPolymorphicBase
    {
        [JsonPropertyName("propA")] public string? PropA { get; init; }
    }

    /// <summary>
    ///     Test derived type B for polymorphic schema testing.
    /// </summary>
    private sealed record TestDerivedB : TestPolymorphicBase
    {
        [JsonPropertyName("propB")] public int PropB { get; init; }
    }

    /// <summary>
    ///     Test result type containing a polymorphic property.
    /// </summary>
    private sealed record TestPolymorphicResult
    {
        [JsonPropertyName("name")] public required string Name { get; init; }

        [JsonPropertyName("details")] public TestPolymorphicBase? Details { get; init; }
    }
    // ReSharper restore UnusedMember.Local
}
