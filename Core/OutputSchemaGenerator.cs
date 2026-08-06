using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using AsposeMcpServer.Results;
using Microsoft.Extensions.AI;

namespace AsposeMcpServer.Core;

/// <summary>
///     Generates JSON Schema for FinalizedResult wrapper containing data and output fields.
/// </summary>
public static class OutputSchemaGenerator
{
    /// <summary>
    ///     Creates a fresh JsonSerializerOptions instance with camelCase naming policy.
    ///     Required because AIJsonUtilities.CreateJsonSchema modifies the options.
    /// </summary>
    /// <returns>A new JsonSerializerOptions with camelCase naming.</returns>
    private static JsonSerializerOptions CreateCamelCaseOptions()
    {
        return new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
    }

    /// <summary>
    ///     Creates schema inference options that force <c>additionalProperties: false</c> on every
    ///     generated object sub-schema. Required to avoid oneOf ambiguity: if one result type's
    ///     required fields are a proper subset of another's (e.g. <c>GetWordContentDetailedResult</c>
    ///     only requires <c>content</c> while <c>GetWordContentResult</c> requires
    ///     <c>{content, totalLength, offset, hasMore}</c>), then a full-content wire payload would
    ///     validate against both branches of the oneOf and Claude Desktop's strict validator would
    ///     report "Tool execution failed" even though the tool ran correctly.
    /// </summary>
    /// <returns>A new options instance with <c>DisallowAdditionalProperties = true</c>.</returns>
    private static AIJsonSchemaCreateOptions CreateSchemaOptions()
    {
        return new AIJsonSchemaCreateOptions
        {
            TransformOptions = new AIJsonSchemaTransformOptions
            {
                DisallowAdditionalProperties = true
            }
        };
    }

    /// <summary>
    ///     Generates JSON Schema from all Handlers in the specified namespace.
    ///     Returns schema for FinalizedResult with data field containing handler result types,
    ///     emitted directly at the top level: the MCP SDK derives the structuredContent wire
    ///     shape from the declared outputSchema, and object-typed schemas are emitted as the
    ///     raw return value on every negotiated protocol version (only non-object schemas get
    ///     the legacy <c>{"result": ...}</c> envelope, and only for pre-SEP-2106 clients).
    /// </summary>
    /// <param name="handlerNamespace">The namespace containing Handler classes.</param>
    /// <param name="assembly">The assembly to scan. Defaults to executing assembly.</param>
    /// <returns>JSON Schema element, or null if no result types found.</returns>
    public static JsonElement? GenerateFromNamespace(string handlerNamespace, Assembly? assembly = null)
    {
        var scanAssembly = assembly ?? Assembly.GetExecutingAssembly();

        var handlerTypes = scanAssembly.GetTypes()
            .Where(t => t.Namespace == handlerNamespace &&
                        t is { IsClass: true, IsAbstract: false } &&
                        t.GetCustomAttribute<ResultTypeAttribute>(true) != null);

        var resultTypes = handlerTypes
            .Select(t => t.GetCustomAttribute<ResultTypeAttribute>(true)!.ResultType)
            .Distinct()
            .ToList();

        if (resultTypes.Count == 0)
            return null;

        return GenerateFinalizedResultSchema(resultTypes);
    }

    /// <summary>
    ///     Generates JSON Schema for a specific type wrapped in FinalizedResult.
    ///     If the type has a static AllTypes property of type Type[], generates oneOf schema for data field.
    /// </summary>
    /// <param name="resultType">The result type to generate schema for.</param>
    /// <returns>JSON Schema element for FinalizedResult wrapper.</returns>
    public static JsonElement GenerateForType(Type resultType)
    {
        var allTypesField = resultType.GetField("AllTypes", BindingFlags.Public | BindingFlags.Static);
        if (allTypesField?.GetValue(null) is Type[] { Length: > 0 } allTypes)
            return GenerateForTypes(allTypes);

        return GenerateFinalizedResultSchema([resultType]);
    }

    /// <summary>
    ///     Generates JSON Schema for multiple types wrapped in FinalizedResult.
    ///     The data field will contain oneOf schema if multiple types provided.
    /// </summary>
    /// <param name="resultTypes">The result types to generate schema for.</param>
    /// <returns>JSON Schema element for FinalizedResult wrapper.</returns>
    public static JsonElement GenerateForTypes(IReadOnlyCollection<Type> resultTypes)
    {
        if (resultTypes.Count == 0)
            throw new ArgumentException("At least one result type is required.", nameof(resultTypes));

        return GenerateFinalizedResultSchema(resultTypes);
    }

    /// <summary>
    ///     Generates the FinalizedResult schema structure with data and output fields.
    /// </summary>
    /// <param name="dataTypes">The possible types for the data field.</param>
    /// <returns>JSON Schema element for FinalizedResult wrapper.</returns>
    private static JsonElement GenerateFinalizedResultSchema(IReadOnlyCollection<Type> dataTypes)
    {
        var outputSchema = AIJsonUtilities.CreateJsonSchema(
            typeof(OutputInfo),
            serializerOptions: CreateCamelCaseOptions(),
            inferenceOptions: CreateSchemaOptions());

        var dataSchema = dataTypes.Count == 1
            ? AIJsonUtilities.CreateJsonSchema(
                dataTypes.First(),
                serializerOptions: CreateCamelCaseOptions(),
                inferenceOptions: CreateSchemaOptions())
            : GenerateOneOfSchema(dataTypes);

        var schemaNode = new JsonObject
        {
            ["type"] = "object",
            ["properties"] = new JsonObject
            {
                ["data"] = JsonNode.Parse(dataSchema.GetRawText()),
                ["output"] = JsonNode.Parse(outputSchema.GetRawText())
            },
            ["required"] = new JsonArray("data", "output"),
            ["additionalProperties"] = false
        };

        var json = schemaNode.ToJsonString();
        return JsonDocument.Parse(json).RootElement.Clone();
    }

    /// <summary>
    ///     Generates oneOf schema for multiple types.
    /// </summary>
    /// <param name="types">The result types.</param>
    /// <returns>oneOf schema element.</returns>
    private static JsonElement GenerateOneOfSchema(IReadOnlyCollection<Type> types)
    {
        var schemas = types.Select(t =>
            AIJsonUtilities.CreateJsonSchema(
                t,
                serializerOptions: CreateCamelCaseOptions(),
                inferenceOptions: CreateSchemaOptions())
        ).ToList();

        var oneOfNode = new JsonObject
        {
            ["oneOf"] = new JsonArray(schemas.Select(s => JsonNode.Parse(s.GetRawText())).ToArray())
        };

        var json = oneOfNode.ToJsonString();
        return JsonDocument.Parse(json).RootElement.Clone();
    }
}
