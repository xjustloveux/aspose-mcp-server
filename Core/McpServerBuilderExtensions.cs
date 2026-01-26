using System.Reflection;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Core;

/// <summary>
///     Extension methods for MCP server builder with tool filtering support.
/// </summary>
public static class McpServerBuilderExtensions
{
    /// <summary>
    ///     Registers tools from the assembly with filtering based on configuration.
    /// </summary>
    /// <param name="builder">The MCP server builder.</param>
    /// <param name="serverConfig">The server configuration for tool filtering.</param>
    /// <param name="sessionConfig">The session configuration for session tool filtering.</param>
    /// <returns>The builder for method chaining.</returns>
    // ReSharper disable once UnusedMethodReturnValue.Global - Fluent API design, return value is optional for chaining
    public static IMcpServerBuilder WithFilteredTools(
        this IMcpServerBuilder builder,
        ServerConfig serverConfig,
        SessionConfig sessionConfig)
    {
        var filterService = new ToolFilterService(serverConfig, sessionConfig);
        var assembly = Assembly.GetExecutingAssembly();

        var toolTypes = assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .ToList();

        foreach (var toolType in toolTypes)
        {
            var methods = toolType.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() != null);

            foreach (var method in methods)
            {
                var toolAttr = method.GetCustomAttribute<McpServerToolAttribute>();
                if (toolAttr != null && !string.IsNullOrEmpty(toolAttr.Name) &&
                    filterService.IsToolEnabled(toolAttr.Name))
                {
                    RegisterToolType(builder, toolType);
                    break;
                }
            }
        }

        return builder;
    }

    /// <summary>
    ///     Registers a tool type using reflection to call the generic WithTools method.
    ///     After registration, applies custom OutputSchema from OutputSchemaAttribute if present.
    /// </summary>
    /// <param name="builder">The MCP server builder.</param>
    /// <param name="toolType">The tool type to register.</param>
    private static void RegisterToolType(IMcpServerBuilder builder, Type toolType)
    {
        var withToolsMethod = typeof(McpServerBuilderExtensions)
            .GetMethod(nameof(RegisterToolGeneric), BindingFlags.NonPublic | BindingFlags.Static)?
            .MakeGenericMethod(toolType);

        withToolsMethod?.Invoke(null, [builder]);
    }

    /// <summary>
    ///     Generic helper method to register a tool type.
    /// </summary>
    /// <typeparam name="T">The tool type to register.</typeparam>
    /// <param name="builder">The MCP server builder.</param>
    private static void RegisterToolGeneric<T>(IMcpServerBuilder builder) where T : class
    {
        builder.WithTools<T>();
    }

    /// <summary>
    ///     Post-processes registered tools to apply custom OutputSchema from OutputSchemaAttribute.
    ///     Call this after all tools are registered.
    /// </summary>
    /// <param name="services">The service collection.</param>
    /// <returns>The service collection for method chaining.</returns>
    public static IServiceCollection ApplyCustomOutputSchemas(this IServiceCollection services)
    {
        return services;
    }

    /// <summary>
    ///     Registers tools with custom OutputSchema support.
    ///     This method creates tools programmatically to allow schema customization.
    ///     Uses deferred service resolution to avoid requiring IServiceProvider at configuration time.
    /// </summary>
    /// <param name="builder">The MCP server builder.</param>
    /// <param name="serverConfig">The server configuration for tool filtering.</param>
    /// <param name="sessionConfig">The session configuration for session tool filtering.</param>
    /// <returns>The builder for method chaining.</returns>
    // ReSharper disable once UnusedMethodReturnValue.Global - Fluent API design, return value is optional for chaining
    public static IMcpServerBuilder WithFilteredToolsAndSchemas(
        this IMcpServerBuilder builder,
        ServerConfig serverConfig,
        SessionConfig sessionConfig)
    {
        var filterService = new ToolFilterService(serverConfig, sessionConfig);
        var assembly = Assembly.GetExecutingAssembly();

        var toolTypes = assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .ToList();

        foreach (var toolType in toolTypes)
        {
            var methods = toolType.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() != null)
                .ToList();

            foreach (var method in methods)
            {
                var toolAttr = method.GetCustomAttribute<McpServerToolAttribute>();
                if (toolAttr == null || string.IsNullOrEmpty(toolAttr.Name) ||
                    !filterService.IsToolEnabled(toolAttr.Name))
                    continue;

                var capturedToolType = toolType;
                var tool = McpServerTool.Create(
                    method,
                    context => ActivatorUtilities.CreateInstance(context.Services!, capturedToolType));

                if (!string.IsNullOrEmpty(tool.ProtocolTool.Description))
                    tool.ProtocolTool.Description =
                        tool.ProtocolTool.Description.Replace("\r\n", "\n").Replace("\r", "\n");

                var mappingAttr = toolType.GetCustomAttribute<ToolHandlerMappingAttribute>();
                if (mappingAttr != null)
                {
                    try
                    {
                        var schema = OutputSchemaGenerator.GenerateFromNamespace(mappingAttr.HandlerNamespace);
                        if (schema.HasValue)
                            tool.ProtocolTool.OutputSchema = schema.Value;
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(
                            $"[WARN] Failed to generate OutputSchema for {toolAttr.Name} from namespace {mappingAttr.HandlerNamespace}: {ex.Message}");
                    }
                }
                else
                {
                    var schemaAttr = method.GetCustomAttribute<OutputSchemaAttribute>();
                    if (schemaAttr != null)
                        try
                        {
                            var schema = OutputSchemaGenerator.GenerateForType(schemaAttr.SchemaType);
                            tool.ProtocolTool.OutputSchema = schema;
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine(
                                $"[WARN] Failed to create OutputSchema for {toolAttr.Name} from {schemaAttr.SchemaType.Name}: {ex.Message}");
                        }
                    else
                        tool.ProtocolTool.OutputSchema = null;
                }

                builder.Services.AddSingleton(tool);
            }
        }

        return builder;
    }
}
