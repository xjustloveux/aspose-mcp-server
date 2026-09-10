using System.Text;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R8-SEC01: what the archive contains, and where each part's references point, comes from the
///     parts the MIME parser produced.
///     <para>
///         The scanner used to read <c>Content-Location</c> out of the raw file with a regex, so a
///         line written inside an HTML body counted as a part header; it gave every part the first
///         remote location it found anywhere as their shared base; it treated a file name as
///         contained whatever host it came from; and a part it could not read came back as an
///         empty string, leaving the archive scanned without it and reported clean.
///     </para>
/// </summary>
[Collection("SerialStaticSeams")]
public class MhtPartModelTests : TestBase
{
    /// <summary>Writes an MHT archive from raw MIME text.</summary>
    /// <param name="name">Fixture file name.</param>
    /// <param name="lines">The archive's lines.</param>
    /// <returns>The archive path.</returns>
    private string Archive(string name, params string[] lines)
    {
        var path = CreateTestFilePath(name);
        File.WriteAllText(path, string.Join("\r\n", lines) + "\r\n", new UTF8Encoding(false));
        return path;
    }

    /// <summary>The refusal an archive produces, or null when it is accepted.</summary>
    /// <param name="path">The archive path.</param>
    /// <returns>The refusal message, or null.</returns>
    private static string? Refusal(string path)
    {
        try
        {
            MhtExternalReferenceScanner.EnsureSelfContained(path);
            return null;
        }
        catch (ArgumentException exception)
        {
            return exception.Message;
        }
    }

    [Fact]
    public void AContentLocationWrittenInsideABody_ShouldNotCountAsAPartHeader()
    {
        // The body claims to carry a part for the image. Only the parser's headers decide that.
        var path = Archive("forged_location.mht",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/index.html", "",
            "<html><body>",
            "Content-Location: http://localhost/tracker.png",
            "<img src=\"tracker.png\"></body></html>", "",
            "--B--");

        var refusal = Refusal(path);

        Assert.NotNull(refusal);
        Assert.Contains("references resources", refusal);
        Assert.Contains("tracker.png", refusal);
    }

    [Fact]
    public void AFileNameCarriedFromOneHost_ShouldNotExcuseTheSameNameOnAnother()
    {
        // logo.png is carried from example.com; the markup's base is another host entirely, so its
        // logo.png is a different file at a different address.
        var path = Archive("basename_reuse.mht",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://first.localhost/index.html", "",
            "<html><body><img src=\"logo.png\"></body></html>", "",
            "--B", "Content-Type: image/png", "Content-Transfer-Encoding: base64",
            "Content-Location: http://second.localhost/logo.png", "",
            "iVBORw0KGgo=", "",
            "--B--");

        var refusal = Refusal(path);

        Assert.NotNull(refusal);
        Assert.Contains("first.localhost/logo.png", refusal);
    }

    // A fixture for "one part's base must not decide another part's references" could not be
    // built through this parser: measured on the pinned Aspose.Email, a second sibling text/html
    // part is exposed as neither an alternate view nor a linked resource — it is dropped, and with
    // it any reference it carried. The shared base was therefore never reachable in practice, and
    // per-part bases are the correct model rather than a fix for a demonstrated escape. What the
    // dropped part does to the *converter* is measured separately, against a listening socket, in
    // MhtUnscannedPartEgressTests: nothing reached the network there.

    [Fact]
    public void APartTheArchiveReallyCarries_ShouldStillBeAccepted()
    {
        // The guard has to leave an ordinary saved page working, relative part location included.
        var path = Archive("contained_ok.mht",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/index.html", "",
            "<html><body><img src=\"logo.png\"></body></html>", "",
            "--B", "Content-Type: image/png", "Content-Transfer-Encoding: base64",
            "Content-Location: logo.png", "",
            "iVBORw0KGgo=", "",
            "--B--");

        Assert.Null(Refusal(path));
    }

    [Fact]
    public void AnArchiveWithNoPartAddressesAtAll_ShouldStillBeReadable()
    {
        // Nothing declares a location, so no relative reference can resolve anywhere: there is
        // nothing to fetch and nothing to refuse.
        var path = Archive("no_locations.mht",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit", "",
            "<html><body><img src=\"logo.png\"></body></html>", "",
            "--B--");

        Assert.Null(Refusal(path));
    }

    [Fact]
    public void APartThatCannotBeRead_ShouldRefuseRatherThanReportClean()
    {
        // The one outcome a fail-closed guard may never produce: a part it could not read used to
        // come back as an empty string, so the archive was scanned without it and declared clean
        // (R8-SEC01). The stream is reached through a seam because nothing in a crafted archive
        // makes the parser's own stream throw.
        var path = Archive("unreadable_part.mht",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/index.html", "",
            "<html><body>nothing remote here</body></html>", "",
            "--B--");

        Assert.Null(Refusal(path));

        var original = MhtExternalReferenceScanner.ReadStreamOf;
        try
        {
            MhtExternalReferenceScanner.ReadStreamOf = _ => throw new IOException("the part is locked");

            var refusal = Refusal(path);

            Assert.NotNull(refusal);
            Assert.Contains("could not be read", refusal, StringComparison.Ordinal);
        }
        finally
        {
            MhtExternalReferenceScanner.ReadStreamOf = original;
        }
    }
}
