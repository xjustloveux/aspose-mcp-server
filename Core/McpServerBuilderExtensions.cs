using System.Reflection;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Core;

/// <summary>
///     Extension methods for MCP server builder with tool filtering support
/// </summary>
public static class McpServerBuilderExtensions
{
    /// <summary>
    ///     Registers tools from the assembly with filtering based on configuration
    /// </summary>
    /// <param name="builder">The MCP server builder</param>
    /// <param name="serverConfig">Server configuration for tool filtering</param>
    /// <param name="sessionConfig">Session configuration for session tool filtering</param>
    /// <returns>The builder for chaining</returns>
    public static IMcpServerBuilder WithFilteredTools(
        this IMcpServerBuilder builder,
        ServerConfig serverConfig,
        SessionConfig sessionConfig)
    {
        var filterService = new ToolFilterService(serverConfig, sessionConfig);
        var assembly = Assembly.GetExecutingAssembly();

        // Find all tool types with McpServerToolType attribute
        var toolTypes = assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .ToList();

        foreach (var toolType in toolTypes)
        {
            // Check if any method in this type should be enabled
            var methods = toolType.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() != null);

            foreach (var method in methods)
            {
                var toolAttr = method.GetCustomAttribute<McpServerToolAttribute>();
                if (toolAttr != null && !string.IsNullOrEmpty(toolAttr.Name))
                    if (filterService.IsToolEnabled(toolAttr.Name))
                    {
                        // Register this tool type if at least one tool is enabled
                        RegisterToolType(builder, toolType);
                        break; // Only register the type once
                    }
            }
        }

        return builder;
    }

    /// <summary>
    ///     Registers a tool type using reflection to call the generic WithTools method
    /// </summary>
    private static void RegisterToolType(IMcpServerBuilder builder, Type toolType)
    {
        // Get the WithTools<T> method
        var withToolsMethod = typeof(McpServerBuilderExtensions)
            .GetMethod(nameof(RegisterToolGeneric), BindingFlags.NonPublic | BindingFlags.Static)?
            .MakeGenericMethod(toolType);

        withToolsMethod?.Invoke(null, [builder]);
    }

    /// <summary>
    ///     Generic helper method to register a tool type
    /// </summary>
    private static void RegisterToolGeneric<T>(IMcpServerBuilder builder) where T : class
    {
        builder.WithTools<T>();
    }
}