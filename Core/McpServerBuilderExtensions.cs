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
        // This method would be called to post-process tools
        // However, accessing ToolCollection requires the McpServer to be built
        // A better approach is to use a hosted service or middleware

        // For now, we'll use a different approach: configure tools during registration
        // See WithFilteredToolsAndSchemas method below
        return services;
    }

    /// <summary>
    ///     Registers tools with custom OutputSchema support.
    ///     This method creates tools programmatically to allow schema customization.
    /// </summary>
    /// <param name="builder">The MCP server builder.</param>
    /// <param name="serverConfig">The server configuration for tool filtering.</param>
    /// <param name="sessionConfig">The session configuration for session tool filtering.</param>
    /// <param name="serviceProvider">Service provider for dependency injection.</param>
    /// <returns>The builder for method chaining.</returns>
    public static IMcpServerBuilder WithFilteredToolsAndSchemas(
        this IMcpServerBuilder builder,
        ServerConfig serverConfig,
        SessionConfig sessionConfig,
        IServiceProvider serviceProvider)
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

                // Create tool with custom options
                var options = new McpServerToolCreateOptions
                {
                    Services = serviceProvider
                };

                // Check for OutputSchemaAttribute
                var schemaAttr = method.GetCustomAttribute<OutputSchemaAttribute>();

                // Create tool instance factory
                var tool = McpServerTool.Create(
                    method,
                    context => ActivatorUtilities.CreateInstance(context.Services!, toolType),
                    options);

                // Apply OutputSchema from ToolHandlerMapping or OutputSchema attribute
                var mappingAttr = toolType.GetCustomAttribute<ToolHandlerMappingAttribute>();
                if (mappingAttr != null)
                    // Use ToolHandlerMapping to generate schema from Handler namespace
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
                else if (schemaAttr != null)
                    // Fallback to OutputSchema attribute for special tools
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

                // Register the tool
                builder.Services.AddSingleton(tool);
            }
        }

        return builder;
    }
}
