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
}
