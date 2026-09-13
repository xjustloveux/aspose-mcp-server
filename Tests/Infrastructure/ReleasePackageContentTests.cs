using System.Text;
using System.Text.RegularExpressions;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Keeps repository tooling and build-only outputs out of the published server.
///     <para>
///         The project uses the Web SDK, which turns every <c>**/*.json</c> under the repository
///         root into a <c>Content</c> item copied to the publish directory. The publish scripts
///         only delete top-level <c>*.json</c>, so <c>graphify/published-docs.json</c> landed in
///         <c>publish/&lt;platform&gt;/graphify/</c> and shipped in every release archive, and
///         the XML documentation file generated for the compiler's CS1591 check shipped beside
///         the single-file executable that nothing can load it for.
///     </para>
/// </summary>
public class ReleasePackageContentTests
{
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

    /// <summary>Reads the production project file.</summary>
    /// <returns>Its full text.</returns>
    private static string ProjectFile()
    {
        return File.ReadAllText(Path.Combine(RepositoryRoot(), "AsposeMcpServer.csproj"), Encoding.UTF8);
    }

    /// <summary>
    ///     A directory that is never a build input must not contribute Content items.
    /// </summary>
    /// <param name="directory">A repository directory holding JSON that the server does not use.</param>
    [Theory]
    [InlineData("docs")]
    [InlineData("graphify")]
    public void NonBuildDirectories_ShouldNotBeCopiedToThePublishOutput(string directory)
    {
        var project = ProjectFile();

        Assert.Matches(new Regex($@"<Content\s+Remove=""{Regex.Escape(directory)}\\\*\*""\s*/>"), project);
    }

    [Fact]
    public void TheDocumentationFile_ShouldNotBePublished()
    {
        var project = ProjectFile();

        Assert.Matches(new Regex(@"<PublishDocumentationFile>\s*false\s*</PublishDocumentationFile>",
            RegexOptions.IgnoreCase), project);
    }
}
