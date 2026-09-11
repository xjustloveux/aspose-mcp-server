using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Structural guard: every Tool that accepts a <see cref="AsposeMcpServer.Core.ServerConfig" /> must hand it to
///     the contexts it builds.
///     <para>
///         The functional tests in <see cref="Tools.ToolAllowlistWiringTests" /> exercise a handful
///         of tools through real dependency injection, which is the only way to prove an allowlist
///         is enforced end to end. What they cannot do is prove the other hundred-odd tools are
///         wired, and sampling is exactly how <c>OcrRecognitionTool</c> and the three OLE tools kept
///         a silent bypass (R2-S01, R2-S02): an empty allowlist means unrestricted, so a tool that
///         never passes its config opts every one of its operations out of <c>--allowed-path</c>
///         without failing anything.
///     </para>
///     <para>
///         This reads the production sources instead, so a new tool that forgets the pass-through
///         fails here even though no functional test covers it yet.
///     </para>
/// </summary>
public class ToolServerConfigWiringTests
{
    private static readonly Regex OperationContextCreation = new("new OperationContext<", RegexOptions.Compiled);

    private static readonly Regex OperationContextWired =
        new(@"ServerConfig\s*=\s*_serverConfig", RegexOptions.Compiled);

    private static readonly Regex DocumentContextCreation =
        new(@"DocumentContext<[^>]*>\s*\.\s*Create\s*\(", RegexOptions.Compiled);

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

    /// <summary>Finds calls to <c>DocumentContext&lt;T&gt;.Create</c> in a source file.</summary>
    /// <param name="source">C# source text.</param>
    /// <returns>The matching invocation expressions.</returns>
    private static IReadOnlyList<InvocationExpressionSyntax> DocumentContextCreations(string source)
    {
        return CSharpSyntaxTree.ParseText(source).GetRoot()
            .DescendantNodes()
            .OfType<InvocationExpressionSyntax>()
            .Where(call => call.Expression is MemberAccessExpressionSyntax
            {
                Expression: GenericNameSyntax { Identifier.ValueText: "DocumentContext" },
                Name.Identifier.ValueText: "Create"
            })
            .ToList();
    }

    /// <summary>Whether a document-context call passes this tool's server configuration.</summary>
    /// <param name="call">The <c>DocumentContext&lt;T&gt;.Create</c> call.</param>
    /// <returns><c>true</c> for the supported named or final positional argument forms.</returns>
    private static bool PassesServerConfig(InvocationExpressionSyntax call)
    {
        var arguments = call.ArgumentList.Arguments;
        if (arguments.Any(argument =>
                argument.NameColon?.Name.Identifier.ValueText == "serverConfig"
                && IsServerConfigField(argument.Expression)))
            return true;

        return arguments.Count > 0
               && arguments[^1].NameColon == null
               && IsServerConfigField(arguments[^1].Expression);
    }

    /// <summary>Whether an expression reads the tool's server-configuration field.</summary>
    /// <param name="expression">The argument expression.</param>
    /// <returns><c>true</c> when it is <c>_serverConfig</c>.</returns>
    private static bool IsServerConfigField(ExpressionSyntax expression)
    {
        return expression is IdentifierNameSyntax { Identifier.ValueText: "_serverConfig" };
    }

    [Fact]
    public void EveryToolThatAcceptsServerConfig_ShouldPassItToTheContextsItBuilds()
    {
        var root = ToolsRoot();
        var offenders = new List<string>();
        var toolsChecked = 0;

        foreach (var file in Directory.GetFiles(root, "*.cs", SearchOption.AllDirectories).OrderBy(f => f))
        {
            var source = File.ReadAllText(file, Encoding.UTF8);
            if (!source.Contains("ServerConfig? serverConfig") && !source.Contains("ServerConfig serverConfig"))
                continue;

            toolsChecked++;
            var name = Path.GetFileName(file);

            var operationContexts = OperationContextCreation.Matches(source).Count;
            var operationContextsWired = OperationContextWired.Matches(source).Count;
            if (operationContexts > operationContextsWired)
                offenders.Add($"{name}: builds {operationContexts} OperationContext but sets " +
                              $"ServerConfig on {operationContextsWired}");

            var documentContextCalls = DocumentContextCreations(source);
            var documentContexts = documentContextCalls.Count;
            var documentContextsWired = documentContextCalls.Count(PassesServerConfig);
            if (documentContexts > documentContextsWired)
                offenders.Add($"{name}: calls DocumentContext.Create {documentContexts} times but " +
                              $"passes serverConfig on {documentContextsWired}");
        }

        Assert.True(toolsChecked > 100,
            $"Only {toolsChecked} tools accept a ServerConfig; the scan is probably looking in the wrong place.");
        Assert.True(offenders.Count == 0,
            "These tools accept a ServerConfig but do not pass it on, so their allowlist is empty " +
            "and therefore unrestricted:" + Environment.NewLine +
            string.Join(Environment.NewLine, offenders));
    }

    [Fact]
    public void EveryToolThatBuildsAContext_ShouldAcceptServerConfig()
    {
        // The scan above only inspects tools that already take a ServerConfig, so a tool whose
        // constructor omits it entirely would never be examined.
        var root = ToolsRoot();
        var offenders = new List<string>();
        var contextBuilders = 0;

        foreach (var file in Directory.GetFiles(root, "*.cs", SearchOption.AllDirectories).OrderBy(f => f))
        {
            var source = File.ReadAllText(file, Encoding.UTF8);
            var buildsContext = OperationContextCreation.IsMatch(source) || DocumentContextCreation.IsMatch(source);
            if (!buildsContext)
                continue;

            contextBuilders++;
            if (!source.Contains("ServerConfig? serverConfig") && !source.Contains("ServerConfig serverConfig"))
                offenders.Add(Path.GetFileName(file));
        }

        // A pattern that stopped matching would leave nothing to inspect and nothing to report,
        // which reads exactly like a clean result (R9-T01).
        Assert.True(contextBuilders > 50,
            $"only {contextBuilders} tools were found to build a context; the patterns this guard "
            + "matches on have probably stopped matching.");

        Assert.True(offenders.Count == 0,
            "These tools build a document or operation context but never accept a ServerConfig, so " +
            "the allowlist can never reach their handlers:" + Environment.NewLine +
            string.Join(Environment.NewLine, offenders));
    }
}
