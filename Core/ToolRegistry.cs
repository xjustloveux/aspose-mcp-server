using System.Reflection;
using AsposeMcpServer.Tools;

namespace AsposeMcpServer.Core;

/// <summary>
///     Tool registry that automatically discovers and registers tools based on naming conventions
/// </summary>
public static class ToolRegistry
{
    /// <summary>
    ///     Discovers and registers all tools based on configuration
    /// </summary>
    /// <param name="config">Server configuration that determines which tool categories to enable</param>
    /// <returns>Dictionary mapping tool names to tool instances</returns>
    public static Dictionary<string, IAsposeTool> DiscoverTools(ServerConfig config)
    {
        var tools = new Dictionary<string, IAsposeTool>();
        var assembly = Assembly.GetExecutingAssembly();

        var toolTypes = assembly.GetTypes()
            .Where(t => typeof(IAsposeTool).IsAssignableFrom(t)
                        && t is { IsInterface: false, IsAbstract: false, Namespace: { } ns } &&
                        ns.StartsWith("AsposeMcpServer.Tools"))
            .ToList();

        foreach (var toolType in toolTypes)
            try
            {
                var tool = (IAsposeTool)Activator.CreateInstance(toolType)!;
                var toolName = GetToolName(toolType);

                if (ShouldRegisterTool(toolName, config))
                    if (!tools.TryAdd(toolName, tool))
                        Console.Error.WriteLine(
                            $"[WARN] Duplicate tool name detected: {toolName}. Skipping {toolType.Name}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[ERROR] Failed to instantiate tool {toolType.Name}: {ex.Message}");
            }

        return tools;
    }

    /// <summary>
    ///     Extracts tool name from type name using naming convention
    ///     Example: WordCreateTool -> word_create
    /// </summary>
    private static string GetToolName(Type toolType)
    {
        var name = toolType.Name;

        if (name.EndsWith("Tool")) name = name.Substring(0, name.Length - 4);

        var snakeCase = string.Concat(name.Select((c, i) =>
            i > 0 && char.IsUpper(c) ? "_" + c.ToString().ToLowerInvariant() : c.ToString().ToLowerInvariant()));

        return snakeCase;
    }

    /// <summary>
    ///     Determines if a tool should be registered based on configuration
    /// </summary>
    private static bool ShouldRegisterTool(string toolName, ServerConfig config)
    {
        if (toolName == "convert_to_pdf") return config.EnableWord || config.EnableExcel || config.EnablePowerPoint;

        if (toolName == "convert_document")
        {
            var enabledDocTools = new[] { config.EnableWord, config.EnableExcel, config.EnablePowerPoint };
            return enabledDocTools.Count(e => e) >= 2;
        }

        if (toolName.StartsWith("word_")) return config.EnableWord;

        if (toolName.StartsWith("excel_")) return config.EnableExcel;

        if (toolName.StartsWith("ppt_")) return config.EnablePowerPoint;

        if (toolName.StartsWith("pdf_")) return config.EnablePdf;

        return true;
    }
}