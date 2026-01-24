using System.Reflection;
using System.Text.Json.Nodes;
using AsposeMcpServer.Core;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tests.Integration.Schema;

/// <summary>
///     Integration tests for validating OutputSchema configuration on all tools.
/// </summary>
[Trait("Category", "Integration")]
public class OutputSchemaValidationTests
{
    private static readonly Assembly TargetAssembly = typeof(OutputSchemaGenerator).Assembly;

    /// <summary>
    ///     Verifies that all tools with ToolHandlerMappingAttribute can generate valid schemas.
    /// </summary>
    [Fact]
    public void AllToolsWithHandlerMapping_ShouldGenerateValidSchema()
    {
        var toolTypes = TargetAssembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null &&
                        t.GetCustomAttribute<ToolHandlerMappingAttribute>() != null)
            .ToList();

        Assert.NotEmpty(toolTypes);

        var failures = new List<string>();

        foreach (var toolType in toolTypes)
        {
            var mappingAttr = toolType.GetCustomAttribute<ToolHandlerMappingAttribute>()!;
            try
            {
                var schema = OutputSchemaGenerator.GenerateFromNamespace(mappingAttr.HandlerNamespace);

                if (schema == null)
                {
                    failures.Add($"{toolType.Name}: Schema is null for namespace {mappingAttr.HandlerNamespace}");
                    continue;
                }

                ValidateSchemaStructure(schema.Value.GetRawText(), toolType.Name, failures);
            }
            catch (Exception ex)
            {
                failures.Add($"{toolType.Name}: Exception - {ex.Message}");
            }
        }

        if (failures.Count > 0)
            Assert.Fail($"Schema validation failures:\n{string.Join("\n", failures)}");
    }

    /// <summary>
    ///     Verifies that all tools with OutputSchemaAttribute can generate valid schemas.
    /// </summary>
    [Fact]
    public void AllToolsWithOutputSchemaAttribute_ShouldGenerateValidSchema()
    {
        var toolMethods = TargetAssembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Instance))
            .Where(m => m.GetCustomAttribute<OutputSchemaAttribute>() != null)
            .ToList();

        if (toolMethods.Count == 0)
            return;

        var failures = new List<string>();

        foreach (var method in toolMethods)
        {
            var schemaAttr = method.GetCustomAttribute<OutputSchemaAttribute>()!;
            var toolName = method.GetCustomAttribute<McpServerToolAttribute>()?.Name ?? method.Name;

            try
            {
                var schema = OutputSchemaGenerator.GenerateForType(schemaAttr.SchemaType);
                ValidateSchemaStructure(schema.GetRawText(), toolName, failures);
            }
            catch (Exception ex)
            {
                failures.Add($"{toolName}: Exception - {ex.Message}");
            }
        }

        if (failures.Count > 0)
            Assert.Fail($"Schema validation failures:\n{string.Join("\n", failures)}");
    }

    /// <summary>
    ///     Verifies that all handler namespaces referenced by tools exist and contain handlers.
    /// </summary>
    [Fact]
    public void AllHandlerNamespaces_ShouldExistAndContainHandlers()
    {
        var toolTypes = TargetAssembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null &&
                        t.GetCustomAttribute<ToolHandlerMappingAttribute>() != null)
            .ToList();

        var failures = new List<string>();

        foreach (var toolType in toolTypes)
        {
            var mappingAttr = toolType.GetCustomAttribute<ToolHandlerMappingAttribute>()!;
            var handlerNamespace = mappingAttr.HandlerNamespace;

            var handlersInNamespace = TargetAssembly.GetTypes()
                .Where(t => t.Namespace == handlerNamespace &&
                            t is { IsClass: true, IsAbstract: false } &&
                            t.GetCustomAttribute<ResultTypeAttribute>(true) != null)
                .ToList();

            if (handlersInNamespace.Count == 0)
                failures.Add($"{toolType.Name}: No handlers with ResultType found in namespace {handlerNamespace}");
        }

        if (failures.Count > 0)
            Assert.Fail($"Handler namespace validation failures:\n{string.Join("\n", failures)}");
    }

    /// <summary>
    ///     Verifies that all handlers have ResultTypeAttribute pointing to valid types.
    /// </summary>
    [Fact]
    public void AllHandlersWithResultType_ShouldHaveValidResultTypes()
    {
        var handlersWithResultType = TargetAssembly.GetTypes()
            .Where(t => t is { IsClass: true, IsAbstract: false } &&
                        t.GetCustomAttribute<ResultTypeAttribute>(true) != null)
            .ToList();

        Assert.NotEmpty(handlersWithResultType);

        var failures = new List<string>();

        foreach (var handler in handlersWithResultType)
        {
            var resultTypeAttr = handler.GetCustomAttribute<ResultTypeAttribute>(true)!;
            var resultType = resultTypeAttr.ResultType;

            if (resultType is not { IsClass: true } and not { IsValueType: true })
            {
                failures.Add($"{handler.Name}: ResultType {resultType.Name} is not a valid type");
                continue;
            }

            try
            {
                var schema = OutputSchemaGenerator.GenerateForType(resultType);
                if (string.IsNullOrEmpty(schema.GetRawText()))
                    failures.Add($"{handler.Name}: Failed to generate schema for {resultType.Name}");
            }
            catch (Exception ex)
            {
                failures.Add($"{handler.Name}: Exception generating schema for {resultType.Name} - {ex.Message}");
            }
        }

        if (failures.Count > 0)
            Assert.Fail($"Handler ResultType validation failures:\n{string.Join("\n", failures)}");
    }

    /// <summary>
    ///     Verifies that the generated schema contains proper JSON Schema structure.
    /// </summary>
    [Fact]
    public void GeneratedSchemas_ShouldBeValidJsonSchema()
    {
        var toolTypes = TargetAssembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null &&
                        t.GetCustomAttribute<ToolHandlerMappingAttribute>() != null)
            .Take(5)
            .ToList();

        foreach (var toolType in toolTypes)
        {
            var mappingAttr = toolType.GetCustomAttribute<ToolHandlerMappingAttribute>()!;
            var schema = OutputSchemaGenerator.GenerateFromNamespace(mappingAttr.HandlerNamespace);

            Assert.NotNull(schema);

            var schemaNode = JsonNode.Parse(schema.Value.GetRawText())!.AsObject();

            Assert.Equal("object", schemaNode["type"]?.GetValue<string>());
            Assert.NotNull(schemaNode["properties"]);
            Assert.NotNull(schemaNode["required"]);
        }
    }

    /// <summary>
    ///     Validates schema structure has required data and output fields.
    /// </summary>
    /// <param name="schemaJson">The schema JSON string.</param>
    /// <param name="toolName">The tool name for error reporting.</param>
    /// <param name="failures">List to collect failures.</param>
    private static void ValidateSchemaStructure(string schemaJson, string toolName, List<string> failures)
    {
        try
        {
            var schemaNode = JsonNode.Parse(schemaJson)!.AsObject();

            if (!schemaNode.ContainsKey("properties"))
            {
                failures.Add($"{toolName}: Schema missing 'properties'");
                return;
            }

            var properties = schemaNode["properties"]!.AsObject();

            if (!properties.ContainsKey("data"))
                failures.Add($"{toolName}: Schema missing 'data' field");

            if (!properties.ContainsKey("output"))
                failures.Add($"{toolName}: Schema missing 'output' field");

            if (schemaNode.ContainsKey("required"))
            {
                var required = schemaNode["required"]!.AsArray();
                var requiredFields = required.Select(r => r?.GetValue<string>()).ToList();

                if (!requiredFields.Contains("data"))
                    failures.Add($"{toolName}: 'data' not in required fields");

                if (!requiredFields.Contains("output"))
                    failures.Add($"{toolName}: 'output' not in required fields");
            }

            if (properties.ContainsKey("output"))
            {
                var outputSchema = properties["output"]!.AsObject();
                if (outputSchema.ContainsKey("properties"))
                {
                    var outputProperties = outputSchema["properties"]!.AsObject();
                    if (!outputProperties.ContainsKey("isSession"))
                        failures.Add($"{toolName}: output schema missing 'isSession' field");
                }
            }
        }
        catch (Exception ex)
        {
            failures.Add($"{toolName}: Failed to parse schema - {ex.Message}");
        }
    }
}
