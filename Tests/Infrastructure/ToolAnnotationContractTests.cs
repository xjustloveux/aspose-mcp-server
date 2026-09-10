using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using AsposeMcpServer.Core;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     The annotations a tool advertises must match what it actually does.
///     <para>
///         <c>readOnlyHint</c> is how a client decides whether a call needs confirmation, so a
///         tool that writes files while claiming to be read-only removes that prompt for exactly
///         the operations that most need it. <c>ocr_recognition</c>, <c>excel_render</c> and
///         <c>word_render</c> all declared <c>ReadOnly = true</c> while writing images and
///         converted documents to caller-supplied paths (R2-S10).
///     </para>
/// </summary>
public class ToolAnnotationContractTests
{
    /// <summary>Parameters through which a caller names somewhere to write.</summary>
    private static readonly Regex DeclaresAnOutputPath = new(
        @"\b(?:outputPath|outputDirectory|outputDir)\b", RegexOptions.Compiled);

    /// <summary>
    ///     Locates the production Tools directory, ignoring the test project's own tree.
    /// </summary>
    /// <returns>The production Tools directory.</returns>
    private static string ToolsRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        var tools = Path.Combine(dir.FullName, "Tools");
        Assert.True(Directory.Exists(tools), $"Production tools not found at {tools}");
        return tools;
    }

    [Fact]
    public void NoToolThatWritesFiles_ShouldAdvertiseItselfAsReadOnly()
    {
        var offenders = new List<string>();
        var toolsChecked = 0;

        foreach (var file in Directory.GetFiles(ToolsRoot(), "*.cs", SearchOption.AllDirectories)
                     .OrderBy(f => f))
        {
            var source = File.ReadAllText(file, Encoding.UTF8);
            var annotation = Regex.Match(source, @"\[McpServerTool\((.*?)\)\]", RegexOptions.Singleline);
            if (!annotation.Success)
                continue;

            toolsChecked++;
            if (!annotation.Groups[1].Value.Contains("ReadOnly = true"))
                continue;

            if (DeclaresAnOutputPath.IsMatch(source))
                offenders.Add(Path.GetFileName(file));
        }

        Assert.True(toolsChecked > 100,
            $"Only {toolsChecked} tools were found; the scan is looking in the wrong place.");
        Assert.True(offenders.Count == 0,
            "These tools declare ReadOnly = true but accept an output path, so a client is told no " +
            "confirmation is needed for a call that writes to disk:" + Environment.NewLine +
            string.Join(Environment.NewLine, offenders));
    }

    [Fact]
    public void EveryRegisteredTool_ShouldCarryAnAnnotationBlock()
    {
        // A tool with no annotations at all inherits the client's defaults, which is a different
        // way of saying nothing about what it does.
        var toolTypes = typeof(ServerConfig).Assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null)
            .ToList();

        Assert.NotEmpty(toolTypes);

        var missing = toolTypes
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() != null)
            .Where(m => string.IsNullOrEmpty(m.GetCustomAttribute<McpServerToolAttribute>()!.Title))
            .Select(m => $"{m.DeclaringType!.Name}.{m.Name}")
            .ToList();

        Assert.True(missing.Count == 0,
            "These tool methods carry no Title, so their annotation block is incomplete:" +
            Environment.NewLine + string.Join(Environment.NewLine, missing));
    }

    [Fact]
    public void EveryToolWithAnArrayParameter_ShouldBoundItAtTheEntryPoint()
    {
        // The transport body cap bounds one request's bytes, not the number of items a single
        // parameter carries, and the handlers only bound inputPaths. Every other array reached
        // its loop unbounded (R2-S07).
        var arrayParameter = new Regex(@"(?:int|string|double)\[\]\??\s+(\w+)\s*=", RegexOptions.Compiled);
        var offenders = new List<string>();

        foreach (var file in Directory.GetFiles(ToolsRoot(), "*.cs", SearchOption.AllDirectories)
                     .OrderBy(f => f))
        {
            var source = File.ReadAllText(file, Encoding.UTF8);
            var parameters = arrayParameter.Matches(source).Select(m => m.Groups[1].Value).Distinct().ToList();
            if (parameters.Count == 0)
                continue;

            var unbounded = new List<string>();
            foreach (var name in parameters)
            {
                var pattern = "ValidateArraySize\\(\\s*" + Regex.Escape(name);
                if (!Regex.IsMatch(source, pattern))
                    unbounded.Add(name);
            }

            if (unbounded.Count > 0)
                offenders.Add($"{Path.GetFileName(file)}: {string.Join(", ", unbounded)}");
        }

        Assert.True(offenders.Count == 0,
            "These tools accept an array without bounding its size at the entry point:" +
            Environment.NewLine + string.Join(Environment.NewLine, offenders));
    }
}
