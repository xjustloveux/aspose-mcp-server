using Aspose.Words.Loading;
using AsposeMcpServer.Helpers;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Guards RB-30: HTML, MHT and content-sniffed Word inputs can reference resources by URI, and
///     loading them turns a conversion into a request issued by the server. The guard must refuse
///     remote URIs while leaving ordinary local image insertion working.
/// </summary>
public class ExternalResourceGuardTests
{
    /// <summary>
    ///     Applies the guard policy to one URI.
    /// </summary>
    /// <param name="uri">The resource URI to test.</param>
    /// <param name="allowed">Allowlist to apply; empty means none configured.</param>
    /// <returns>The action the guard would take.</returns>
    private static ResourceLoadingAction Decide(string uri, params string[] allowed)
    {
        return ExternalResourceGuard.IsAllowed(uri, allowed)
            ? ResourceLoadingAction.Default
            : ResourceLoadingAction.Skip;
    }

    [Theory]
    [InlineData("http://169.254.169.254/latest/meta-data/")]
    [InlineData("https://internal.example.com/logo.png")]
    [InlineData("ftp://example.com/a.png")]
    public void RemoteUri_ShouldBeSkipped(string uri)
    {
        Assert.Equal(ResourceLoadingAction.Skip, Decide(uri));
    }

    [Fact]
    public void DataUri_ShouldBeLoaded()
    {
        Assert.Equal(ResourceLoadingAction.Default, Decide("data:image/png;base64,AAAA"));
    }

    [Fact]
    public void LocalAbsolutePath_WithNoAllowlist_ShouldBeLoaded()
    {
        var path = Path.Combine(Path.GetTempPath(), "guard-fixture.png");

        Assert.Equal(ResourceLoadingAction.Default, Decide(path));
    }

    [Fact]
    public void LocalRelativePath_WithNoAllowlist_ShouldBeLoaded()
    {
        Assert.Equal(ResourceLoadingAction.Default, Decide("images/logo.png"));
    }

    [Fact]
    public void LocalPathInsideAllowlist_ShouldBeLoaded()
    {
        var root = Path.Combine(Path.GetTempPath(), "GuardAllowed_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var inside = Path.Combine(root, "logo.png");
            File.WriteAllBytes(inside, [0x89, 0x50]);

            Assert.Equal(ResourceLoadingAction.Default, Decide(inside, root));
        }
        finally
        {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void LocalPathOutsideAllowlist_ShouldBeSkipped()
    {
        var root = Path.Combine(Path.GetTempPath(), "GuardAllowed_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            var outside = Path.Combine(Path.GetTempPath(), "elsewhere.png");

            Assert.Equal(ResourceLoadingAction.Skip, Decide(outside, root));
        }
        finally
        {
            Directory.Delete(root, true);
        }
    }

    /// <summary>
    ///     A network location must be refused whether or not an allowlist is configured.
    ///     <para>
    ///         "No allowlist means local references are permitted" was implemented as "anything
    ///         .NET calls a file URI is permitted", and .NET reports <c>file://host/share</c> as
    ///         <c>IsFile = true</c>. Loading one makes the server open an SMB connection to a host
    ///         the document names, which leaks NTLM credentials and fetches remote content
    ///         (R2-S05). A protocol-relative reference is not even an absolute URI, so it skipped
    ///         the remote check altogether.
    ///     </para>
    /// </summary>
    /// <param name="uri">A reference that resolves to another host.</param>
    [Theory]
    [InlineData("file://attacker.example.com/share/a.png")]
    [InlineData("file://attacker/share/a.png")]
    [InlineData("\\\\attacker\\share\\a.png")]
    [InlineData("//attacker/share/a.png")]
    public void NetworkLocation_WithoutAllowlist_ShouldBeSkipped(string uri)
    {
        Assert.Equal(ResourceLoadingAction.Skip, Decide(uri));
    }

    [Theory]
    [InlineData("file://attacker.example.com/share/a.png")]
    [InlineData("\\\\attacker\\share\\a.png")]
    [InlineData("//attacker/share/a.png")]
    public void NetworkLocation_WithAllowlist_ShouldBeSkipped(string uri)
    {
        var root = Path.Combine(Path.GetTempPath(), "GuardUnc_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            Assert.Equal(ResourceLoadingAction.Skip, Decide(uri, root));
        }
        finally
        {
            Directory.Delete(root, true);
        }
    }

    [Theory]
    [InlineData("file:///C:/local/a.png")]
    [InlineData("/tmp/local/a.png")]
    public void LocalReference_WithoutAllowlist_ShouldStillBeLoaded(string uri)
    {
        // The documented behaviour for an unconfigured allowlist is unchanged; only references
        // that name another host are newly refused.
        Assert.Equal(ResourceLoadingAction.Default, Decide(uri));
    }

    [Fact]
    public void EmptyUri_ShouldBeLoaded()

    {
        Assert.Equal(ResourceLoadingAction.Default, Decide(string.Empty));
    }
}
