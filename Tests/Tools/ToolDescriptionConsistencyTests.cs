using System.ComponentModel;
using System.Reflection;
using System.Text.RegularExpressions;
using AsposeMcpServer.Core;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tests.Tools;

/// <summary>
///     Guards the operation lists in tool <see cref="DescriptionAttribute" /> texts against the
///     actual handler dispatch keys. Operation renames that only touch the handlers silently
///     leave the user-facing descriptions advertising operations that no longer exist (or hiding
///     ones that do); this test turns that drift into a build-time failure with the full list.
/// </summary>
public class ToolDescriptionConsistencyTests
{
    private static readonly Regex OperationListPattern = new(
        @"Supports\s+(\d+)\s+operations?:\s*([^.\r\n]+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     Tool-layer operation aliases: valid wire operations that the tool itself remaps to a
    ///     handler dispatch key (e.g. word_field maps update_all → update plus a flag), so they
    ///     legitimately appear in the description without a handler of their own.
    /// </summary>
    private static readonly Dictionary<string, string[]> ToolLayerAliases = new()
    {
        ["WordFieldTool"] = ["update_all"]
    };

    [Fact]
    public void ToolDescriptions_OperationLists_MatchHandlerDispatchKeys()
    {
        var assembly = typeof(ToolHandlerMappingAttribute).Assembly;
        var failures = new List<string>();
        var checkedTools = 0;

        foreach (var toolType in assembly.GetTypes())
        {
            var mapping = toolType.GetCustomAttribute<ToolHandlerMappingAttribute>();
            if (mapping == null) continue;

            var executeMethod = toolType.GetMethods()
                .FirstOrDefault(m => m.GetCustomAttribute<McpServerToolAttribute>() != null);
            var description = executeMethod?.GetCustomAttribute<DescriptionAttribute>()?.Description;
            if (description == null) continue;

            var match = OperationListPattern.Match(description);
            if (!match.Success) continue;

            var declaredCount = int.Parse(match.Groups[1].Value);
            var declaredOps = match.Groups[2].Value
                .Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
                .Select(o => o.ToLowerInvariant())
                .ToHashSet();

            var actualOps = GetHandlerOperations(assembly, mapping.HandlerNamespace);
            if (actualOps.Count == 0) continue;

            checkedTools++;
            var aliases = ToolLayerAliases.GetValueOrDefault(toolType.Name, []);
            var ghosts = declaredOps.Except(actualOps).Except(aliases).OrderBy(x => x).ToList();
            var missing = actualOps.Except(declaredOps).OrderBy(x => x).ToList();

            if (ghosts.Count > 0 || missing.Count > 0 || declaredCount != declaredOps.Count)
                failures.Add(
                    $"{toolType.Name}: description declares {declaredCount} operations " +
                    $"[{string.Join(", ", declaredOps.OrderBy(x => x))}] but handlers implement " +
                    $"[{string.Join(", ", actualOps.OrderBy(x => x))}]" +
                    (ghosts.Count > 0 ? $" — ghost operations: {string.Join(", ", ghosts)}" : "") +
                    (missing.Count > 0 ? $" — undocumented operations: {string.Join(", ", missing)}" : ""));
        }

        Assert.True(checkedTools > 0, "No tool descriptions with an operation list were found — pattern broken?");
        Assert.True(failures.Count == 0,
            "Tool [Description] operation lists must match handler dispatch keys:\n" + string.Join("\n", failures));
    }

    /// <summary>
    ///     Collects the dispatch keys of every concrete handler in the given namespace.
    /// </summary>
    /// <param name="assembly">The assembly containing the handlers.</param>
    /// <param name="handlerNamespace">The handler namespace declared by the tool.</param>
    /// <returns>The lower-cased operation keys.</returns>
    private static HashSet<string> GetHandlerOperations(Assembly assembly, string handlerNamespace)
    {
        return assembly.GetTypes()
            .Where(t => t.Namespace == handlerNamespace && t is { IsAbstract: false, IsInterface: false })
            .Select(t => (Type: t, OperationProperty: t.GetProperty("Operation")))
            .Where(x => x.OperationProperty?.PropertyType == typeof(string) &&
                        x.Type.GetConstructor(Type.EmptyTypes) != null)
            .Select(x => ((string)x.OperationProperty!.GetValue(Activator.CreateInstance(x.Type))!).ToLowerInvariant())
            .ToHashSet();
    }
}
