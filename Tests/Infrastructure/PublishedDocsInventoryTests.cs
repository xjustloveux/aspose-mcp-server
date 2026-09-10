using System.Text.Json;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Nothing may sit under <c>docs/</c> unless the publish manifest names it, and the workflow
///     that publishes must consult that manifest before it uploads.
///     <para>
///         <c>deploy-pages.yml</c> uploads <c>./docs</c> whole, so every file under it goes live.
///         <c>verify-public-map.ps1</c> says exactly that and acts on it — for
///         <c>docs/architecture-map/</c> only. The first version of these rules lived here, in a
///         test the deploy workflow does not run, so a file placed anywhere else under
///         <c>docs/</c> was published unlisted regardless of what this class asserted (R13-G01).
///     </para>
///     <para>
///         The manifest now lives in <c>graphify/published-docs.json</c> and the gate in
///         <c>graphify/verify-published-docs.ps1</c>, which the workflow runs before the upload.
///         This class reads the same file — so the two cannot drift — and exercises the rules with
///         the negative cases a CI script cannot carry.
///     </para>
/// </summary>
public class PublishedDocsInventoryTests : TestBase
{
    private static readonly JsonSerializerOptions ManifestJsonOptions = new()
        { PropertyNameCaseInsensitive = true };

    /// <summary>Walks up from the test binaries to the repository root.</summary>
    /// <returns>The repository root directory.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }

    /// <summary>Reads the manifest the deploy gate reads.</summary>
    /// <returns>What may be published, and what is pending removal.</returns>
    private static Manifest ReadManifest()
    {
        var path = Path.Combine(RepositoryRoot().FullName, "graphify", "published-docs.json");
        Assert.True(File.Exists(path), "the publish manifest is missing");

        var manifest = JsonSerializer.Deserialize<Manifest>(File.ReadAllText(path),
            ManifestJsonOptions);

        Assert.NotNull(manifest);
        return manifest;
    }

    /// <summary>Every file that currently lives under docs/.</summary>
    /// <returns>Repository-relative paths, in posix form.</returns>
    private static List<string> FilesUnderDocs()
    {
        var root = RepositoryRoot();
        var docs = Path.Combine(root.FullName, "docs");

        return Directory.EnumerateFiles(docs, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(root.FullName, path).Replace('\\', '/'))
            .OrderBy(path => path, StringComparer.Ordinal)
            .ToList();
    }

    [Fact]
    public void NothingUnderDocs_ShouldBeUnaccountedFor()
    {
        var manifest = ReadManifest();
        var accounted = manifest.Published
            .Concat(manifest.PendingRemoval.Select(entry => entry.Path))
            .ToHashSet(StringComparer.Ordinal);

        var unexpected = FilesUnderDocs()
            .Where(path => !accounted.Contains(path))
            .ToList();

        Assert.True(unexpected.Count == 0,
            "these files are under docs/ and therefore published, and the manifest does not "
            + "account for them: " + string.Join(", ", unexpected));
    }

    [Fact]
    public void EveryPublishedEntry_ShouldStillExist()
    {
        // A manifest that outlives its files stops describing anything, and a missing page leaves
        // the sitemap pointing at nothing.
        var present = FilesUnderDocs().ToHashSet(StringComparer.Ordinal);

        var missing = ReadManifest().Published.Where(path => !present.Contains(path)).ToList();

        Assert.True(missing.Count == 0,
            "these are listed as published but are not there: " + string.Join(", ", missing));
    }

    [Fact]
    public void EveryPendingRemoval_ShouldStillBePending()
    {
        // When the files go, the entry goes with them. Leaving it behind quietly permits their
        // return.
        var present = FilesUnderDocs().ToHashSet(StringComparer.Ordinal);

        var gone = ReadManifest().PendingRemoval
            .Where(entry => !present.Contains(entry.Path))
            .Select(entry => entry.Path)
            .ToList();

        Assert.True(gone.Count == 0,
            "these are recorded as pending removal but are already gone; delete the entry too: "
            + string.Join(", ", gone));
    }

    [Fact]
    public void EveryPendingRemoval_ShouldSayWhatHasToHappen()
    {
        // "Known but unresolved" is only a record if it records the resolution.
        var vague = ReadManifest().PendingRemoval
            .Where(entry => string.IsNullOrWhiteSpace(entry.Disposition))
            .Select(entry => entry.Path)
            .ToList();

        Assert.True(vague.Count == 0,
            "these pending-removal entries say nothing about what has to happen to them: "
            + string.Join(", ", vague));
    }

    [Fact]
    public void TheDeployWorkflow_ShouldRunTheManifestGateBeforeUploading()
    {
        // The whole point of R13-G01: rules that the publishing path does not run are not rules.
        var workflow = Path.Combine(RepositoryRoot().FullName,
            ".github", "workflows", "deploy-pages.yml");
        Assert.True(File.Exists(workflow), "the deploy workflow is missing");

        var text = File.ReadAllText(workflow);

        var gate = text.IndexOf("verify-published-docs.ps1", StringComparison.Ordinal);
        var upload = text.IndexOf("upload-pages-artifact", StringComparison.Ordinal);

        // Compared against the value IndexOf actually returns when it finds nothing. `>= 0` says
        // the same thing less directly, and reads as the vacuous bound it usually is.
        Assert.True(gate != -1,
            "the deploy workflow does not run the docs manifest gate, so the manifest governs "
            + "nothing that actually gets published");
        Assert.True(upload != -1, "the deploy workflow no longer uploads an artifact");
        Assert.True(gate < upload,
            "the manifest gate runs after the upload, so an unlisted file would already be live");
        // Comment lines skipped: the workflow explains *why* there is no continue-on-error, and
        // a plain substring match found that sentence. The same mistake this round keeps
        // correcting elsewhere, made here in the check for it.
        var directives = text.Split('\n')
            .Where(line => !line.TrimStart().StartsWith('#'))
            .ToList();

        // Counted, and the split checked first. A predicate DoesNotContain over a collection is
        // satisfied by an empty collection, so a split that read nothing would have passed this
        // while checking nothing.
        Assert.True(directives.Count > 10,
            $"the workflow read as {directives.Count} directive lines, so nothing was checked");
        Assert.Equal(0, directives.Count(line =>
            line.Contains("continue-on-error", StringComparison.Ordinal)));
    }

    [Fact]
    public void TheScan_ShouldFindTheSiteItself()
    {
        // A scan reading nothing would make every assertion above vacuous.
        var files = FilesUnderDocs();
        var manifest = ReadManifest();

        Assert.Contains("docs/index.html", files);
        Assert.True(files.Count >= manifest.Published.Count,
            $"only {files.Count} files under docs/; the manifest names {manifest.Published.Count}");
    }

    /// <summary>The manifest's shape.</summary>
    /// <param name="Published">Files that may go live.</param>
    /// <param name="PendingRemoval">Files present but not meant to be published.</param>
    private sealed record Manifest(
        List<string> Published,
        List<PendingEntry> PendingRemoval);

    /// <summary>A file that is present but should not be published.</summary>
    /// <param name="Path">Its repository-relative path.</param>
    /// <param name="Disposition">What has to happen to it.</param>
    // ReSharper disable once ClassNeverInstantiated.Local -- System.Text.Json constructs manifest entries.
    private sealed record PendingEntry(string Path, string Disposition);
}
