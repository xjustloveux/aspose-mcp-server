using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using AsposeMcpServer.Core;
using AsposeMcpServer.Results.Common;

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
}
