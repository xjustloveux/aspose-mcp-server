using System.Text;
using System.Text.RegularExpressions;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     The configuration reference must list every switch the parser understands.
///     <para>
///         A security setting that exists but is not documented is one nobody configures:
///         <c>--auth-apikey-trusted-proxies</c>, <c>--auth-jwt-trusted-proxies</c> and the metrics
///         authentication switches were all absent from the reference page while being the
///         settings that decide whether identity headers are trusted and whether the metrics
///         endpoint is public (R2-C06).
///     </para>
///     <para>
///         The inventory is taken from the parser rather than from a hand-maintained list, so a
///         new switch is undocumented by default and has to be added deliberately.
///     </para>
/// </summary>
public class ConfigurationDocumentationTests
{
    /// <summary>
    ///     Switches that are deliberately not in the configuration reference, each with the reason.
    ///     Keeping them here rather than loosening the check means the list itself is reviewable.
    /// </summary>
    private static readonly Dictionary<string, string> NotInTheReference = new()
    {
        ["--powerpoint"] = "Long-form alias of --ppt, which the reference documents.",
        ["--yes"] = "Not a server switch: an argument this server passes to npx when resolving an "
                    + "extension command. It is picked up by the inventory scan, not by the parser.",
        ["--extension-frame-interval-default"] = "Alias of --extension-frame-interval, documented.",
        ["--extension-snapshot-ttl-default"] = "Alias of --extension-snapshot-ttl, documented.",
        ["--extension-max-missed-heartbeats-default"] =
            "Alias of --extension-max-missed-heartbeats, documented.",
        ["--extension-idle-timeout-default"] = "Alias of --extension-idle-timeout, documented."
    };

    /// <summary>Aliases whose documented counterpart the exemption above claims exists.</summary>
    private static readonly Dictionary<string, string> AliasOf = new()
    {
        ["--powerpoint"] = "--ppt",
        ["--extension-frame-interval-default"] = "--extension-frame-interval",
        ["--extension-snapshot-ttl-default"] = "--extension-snapshot-ttl",
        ["--extension-max-missed-heartbeats-default"] = "--extension-max-missed-heartbeats",
        ["--extension-idle-timeout-default"] = "--extension-idle-timeout"
    };

    /// <summary>
    ///     Settings a deployment has to decide about deliberately, so the example must show how
    ///     each is written rather than leaving an operator to find it in the reference table.
    /// </summary>
    private static readonly string[] MustAppearInTheExample =
    [
        "--auth-apikey-trusted-proxies",
        "--auth-jwt-trusted-proxies",
        "--allow-external-resources",
        "--metrics-require-auth",
        "--metrics-path",
        "--extension-max-restarts",
        "--extension-max-missed-heartbeats",
        "--extension-idle-timeout",
        "--extension-snapshot-ttl"
    ];

    /// <summary>
    ///     Locates the repository root by walking up to the production project file.
    /// </summary>
    /// <returns>The repository root directory.</returns>
    private static string RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir.FullName;
    }

    /// <summary>
    ///     Collects every command-line switch and environment variable the parser recognises.
    /// </summary>
    /// <param name="root">Repository root.</param>
    /// <returns>The switches and the environment variable names.</returns>
    private static (HashSet<string> Flags, HashSet<string> Environment) ParserInventory(string root)
    {
        var flags = new HashSet<string>(StringComparer.Ordinal);
        var environment = new HashSet<string>(StringComparer.Ordinal);

        var sources = Directory.GetFiles(Path.Combine(root, "Core"), "*.cs", SearchOption.AllDirectories)
            .Append(Path.Combine(root, "Program.cs"))
            .Where(File.Exists);

        foreach (var file in sources)
        {
            var text = File.ReadAllText(file, Encoding.UTF8);
            foreach (Match m in Regex.Matches(text, "\"(--[a-z0-9-]+)\""))
                flags.Add(m.Groups[1].Value);
            foreach (Match m in Regex.Matches(text, "\"(ASPOSE_[A-Z0-9_]+)\""))
                environment.Add(m.Groups[1].Value);
        }

        return (flags, environment);
    }

    [Fact]
    public void EverySwitchTheParserAccepts_ShouldAppearInTheConfigurationReference()
    {
        var root = RepositoryRoot();
        var reference = File.ReadAllText(Path.Combine(root, "docs", "configuration.html"), Encoding.UTF8);
        var (flags, environment) = ParserInventory(root);

        Assert.True(flags.Count > 50,
            $"Only {flags.Count} switches were found; the scan is looking in the wrong place.");

        var undocumented = flags
            .Where(f => !NotInTheReference.ContainsKey(f))
            .Where(f => !reference.Contains(f, StringComparison.Ordinal))
            .OrderBy(f => f, StringComparer.Ordinal)
            .Concat(environment
                .Where(e => !reference.Contains(e, StringComparison.Ordinal))
                .OrderBy(e => e, StringComparer.Ordinal))
            .ToList();

        Assert.True(undocumented.Count == 0,
            "These settings are accepted by the parser but absent from docs/configuration.html:" +
            Environment.NewLine + string.Join(Environment.NewLine, undocumented));
    }

    [Fact]
    public void TheExemptionList_ShouldNotOutliveTheSwitchesItCovers()
    {
        // An exemption for a switch that no longer exists hides the fact that the list is stale.
        var (flags, _) = ParserInventory(RepositoryRoot());

        var stale = NotInTheReference.Keys.Where(f => !flags.Contains(f)).OrderBy(f => f).ToList();

        Assert.True(stale.Count == 0,
            "These exemptions name switches the parser no longer accepts:" +
            Environment.NewLine + string.Join(Environment.NewLine, stale));
    }

    [Fact]
    public void EveryAliasExemption_ShouldNameASwitchThatIsActuallyDocumented()
    {
        // An exemption is only honest if its stated reason holds. The first version of this list
        // claimed nine extension switches were "documented on the extensions page"; none of them
        // were, and the guard passed anyway because nothing checked the reason.
        var root = RepositoryRoot();
        var reference = File.ReadAllText(Path.Combine(root, "docs", "configuration.html"), Encoding.UTF8);
        var (flags, _) = ParserInventory(root);

        var broken = new List<string>();
        foreach (var (alias, documented) in AliasOf)
            if (!flags.Contains(alias))
                broken.Add($"{alias}: the parser no longer accepts it");
            else if (!flags.Contains(documented))
                broken.Add($"{alias}: claims to alias {documented}, which the parser does not accept");
            else if (!reference.Contains(documented, StringComparison.Ordinal))
                broken.Add($"{alias}: claims {documented} is documented, but it is not");

        Assert.True(broken.Count == 0,
            "These alias exemptions do not hold:" + Environment.NewLine + string.Join(Environment.NewLine, broken));
    }

    /// <summary>
    ///     Every switch the example shows must be one the parser accepts. A hand-maintained
    ///     example drifts silently: nothing fails when it names an option that no longer exists,
    ///     and an operator copying it gets a server that ignores the line (R3-DOC08).
    /// </summary>
    [Fact]
    public void EverySwitchInTheExample_ShouldBeOneTheParserAccepts()
    {
        var root = RepositoryRoot();
        var (flags, _) = ParserInventory(root);
        var example = File.ReadAllText(Path.Combine(root, "config_example.json"), Encoding.UTF8);

        var used = Regex.Matches(example, "\"(--[a-z0-9-]+)\"")
            .Select(m => m.Groups[1].Value)
            .Distinct()
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();

        Assert.NotEmpty(used);

        var unknown = used.Where(f => !flags.Contains(f)).ToList();
        Assert.True(unknown.Count == 0,
            "config_example.json names switches the parser does not accept:" + Environment.NewLine +
            string.Join(Environment.NewLine, unknown));
    }

    /// <summary>
    ///     The example showed neither the trusted-proxy settings, the external-resource policy,
    ///     the metrics authentication settings nor the extension resilience settings, which are
    ///     exactly the ones a deployment has to think about (R3-DOC08).
    /// </summary>
    [Fact]
    public void TheExample_ShouldShowTheSettingsADeploymentHasToDecide()
    {
        var example = File.ReadAllText(Path.Combine(RepositoryRoot(), "config_example.json"),
            Encoding.UTF8);

        var absent = MustAppearInTheExample
            .Where(flag => !example.Contains("\"" + flag + "\"", StringComparison.Ordinal))
            .ToList();

        Assert.True(absent.Count == 0,
            "config_example.json does not show these settings:" + Environment.NewLine +
            string.Join(Environment.NewLine, absent));
    }

    /// <summary>
    ///     A required setting that the parser no longer accepts would make the check above demand
    ///     something impossible, so the list is held to the same standard as the example.
    /// </summary>
    [Fact]
    public void TheRequiredExampleSettings_ShouldAllStillExist()
    {
        var (flags, _) = ParserInventory(RepositoryRoot());

        var gone = MustAppearInTheExample.Where(flag => !flags.Contains(flag)).ToList();

        Assert.True(gone.Count == 0,
            "These settings are required in the example but the parser no longer accepts them:" +
            Environment.NewLine + string.Join(Environment.NewLine, gone));
    }
}
