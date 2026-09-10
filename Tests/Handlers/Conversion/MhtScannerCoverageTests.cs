using System.Text;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     Covers A-03's scanner half. The guard refuses an MHT archive that would make the server
///     fetch something, but it only recognised references written as <c>scheme://host</c>. A
///     protocol-relative reference fetches just the same, <c>file:</c> was treated as harmless
///     without checking where it points, and a file the parser could not read was allowed through
///     on the assumption that the converter would reject it — while a different parser inside the
///     converter may well accept it.
/// </summary>
public class MhtScannerCoverageTests : TestBase
{
    /// <summary>Writes an MHT archive whose single HTML part holds the supplied markup.</summary>
    /// <param name="fileName">File name under the test directory.</param>
    /// <param name="html">Markup for the HTML part.</param>
    /// <returns>The path written.</returns>
    private string WriteMht(string fileName, string html)
    {
        var path = CreateTestFilePath(fileName);
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "Subject: fixture",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"----=_Boundary\"",
            "",
            "------=_Boundary",
            "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/fixture.html",
            "",
            html,
            "",
            "------=_Boundary--",
            ""), Encoding.UTF8);
        return path;
    }

    /// <summary>
    ///     A <c>file:</c> reference written without the double slash must still be scanned.
    ///     <para>
    ///         The pattern captures that shape in its own <c>uri4</c> group, but the reader only
    ///         looked at <c>uri</c>, <c>uri2</c> and <c>uri3</c>, so the match came back as an
    ///         empty string and the loop skipped it as if nothing had matched (R2-S04). The archive
    ///         was then judged self-contained and handed to the Aspose MHT loader with a local
    ///         path in it.
    ///     </para>
    /// </summary>
    /// <param name="uri">A file reference the scanner has to notice.</param>
    [Theory]
    [InlineData("file:/C:/Windows/win.ini")]
    [InlineData("file:/etc/passwd")]
    [InlineData("file:///C:/Windows/win.ini")]
    [InlineData("file://attacker.example.com/share/payload.png")]
    public void FileReferenceWithoutAllowlist_ShouldBeReported(string uri)
    {
        var path = WriteMht("file_ref.mht", $"<html><body><img src=\"{uri}\"></body></html>");

        Assert.NotEmpty(MhtExternalReferenceScanner.FindRemoteReferences(path));
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    /// <summary>
    ///     A single-slash <c>file:</c> reference must be noticed in CSS contexts too, not only in
    ///     element attributes. <c>url()</c> and <c>@import</c> both fetch.
    /// </summary>
    /// <param name="markup">Body markup carrying the reference.</param>
    [Theory]
    [InlineData("<style>body{background:url(file:/C:/Windows/win.ini)}</style>")]
    [InlineData("<style>body{background:url(\"file:/etc/passwd\")}</style>")]
    [InlineData("<style>@import \"file:/etc/passwd\";</style>")]
    public void FileReferenceInCssWithoutAllowlist_ShouldBeReported(string markup)
    {
        var path = WriteMht("file_css.mht", $"<html><body>{markup}</body></html>");

        Assert.NotEmpty(MhtExternalReferenceScanner.FindRemoteReferences(path));
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    /// <summary>
    ///     A relative reference is remote when the archive says its base is remote.
    ///     <para>
    ///         The fixture helper declares <c>Content-Location: http://localhost/fixture.html</c>,
    ///         which is what a saved web page carries. A relative <c>src</c> normally names another
    ///         part of the archive; when no such part exists it resolves against that base instead.
    ///         Measured with a localhost probe server against the pinned Aspose.Pdf: the converter
    ///         resolves it and fetches it, and the pattern-based scan could not see it because the
    ///         reference carries no scheme and the base lives in a MIME header (§13.7 item 1).
    ///     </para>
    /// </summary>
    [Fact]
    public void RelativeReferenceAgainstARemoteBase_ShouldBeReported()
    {
        var path = WriteMht("relative_remote_base.mht",
            "<html><body><img src=\"relative-probe.png\"></body></html>");

        Assert.NotEmpty(MhtExternalReferenceScanner.FindRemoteReferences(path));
        var exception = Assert.ThrowsAny<ArgumentException>(() =>
            MhtExternalReferenceScanner.EnsureSelfContained(path));
        Assert.Contains("references resources", exception.Message);
    }

    [Fact]
    public void RelativeReferenceNamingAnotherPart_ShouldBeAccepted()
    {
        // The ordinary shape of a saved page: the image is a second part of the archive, and the
        // markup names it relatively. Nothing has to be fetched, so nothing is refused.
        var path = CreateTestFilePath("relative_contained.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <saved by test>",
            "Subject: fixture",
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"----=_Boundary\"",
            "",
            "------=_Boundary",
            "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/fixture.html",
            "",
            "<html><body><img src=\"relative-probe.png\"></body></html>",
            "",
            "------=_Boundary",
            "Content-Type: image/png",
            "Content-Transfer-Encoding: base64",
            "Content-Location: relative-probe.png",
            "",
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
            "",
            "------=_Boundary--",
            ""), Encoding.UTF8);

        var exception = Record.Exception(() =>
            MhtExternalReferenceScanner.EnsureSelfContained(path));

        Assert.Null(exception);
    }

    /// <summary>
    ///     Every relative syntax the pinned loader actually fetches must be scanned.
    ///     <para>
    ///         Measured with a localhost probe: an unquoted attribute value, a relative
    ///         <c>url()</c> and a relative <c>@import</c> were each fetched when the archive
    ///         declared a remote base. A relative <c>srcset</c> candidate was not — only the
    ///         accompanying <c>src</c> — so srcset is deliberately not matched (R3-S05).
    ///     </para>
    /// </summary>
    /// <param name="markup">Body markup carrying the relative reference.</param>
    [Theory]
    [InlineData("<img src=rel-unquoted.png >")]
    [InlineData("<style>body{background:url(rel-cssurl.png)}</style>")]
    [InlineData("<style>@import \"rel-cssimport.css\";</style>")]
    public void RelativeReferenceInAnySupportedSyntax_ShouldBeReported(string markup)
    {
        var path = WriteMht("relative_syntax.mht", "<html><body>" + markup + "</body></html>");

        Assert.NotEmpty(MhtExternalReferenceScanner.FindRemoteReferences(path));
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Theory]
    [InlineData("//evil.example.com/tracker.png")]
    [InlineData("//evil.example.com:8080/tracker.png")]
    public void ProtocolRelativeReference_ShouldBeReported(string uri)
    {
        // The browser or renderer supplies the scheme and fetches it; leaving the scheme out is
        // not a way of staying local.
        var path = WriteMht("protocol_relative.mht", $"<html><body><img src=\"{uri}\"></body></html>");

        Assert.NotEmpty(MhtExternalReferenceScanner.FindRemoteReferences(path));
        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void FileReferenceOutsideTheAllowlist_ShouldBeReported()
    {
        var allowed = Path.Combine(TestDir, "allowed_root");
        Directory.CreateDirectory(allowed);
        var outside = Path.Combine(TestDir, "outside_root", "secret.txt");
        Directory.CreateDirectory(Path.GetDirectoryName(outside)!);
        File.WriteAllText(outside, "secret");

        var uri = "file:///" + outside.Replace("\\", "/");
        var path = WriteMht("file_outside.mht", $"<html><body><img src=\"{uri}\"></body></html>");

        Assert.ThrowsAny<ArgumentException>(() =>
            MhtExternalReferenceScanner.EnsureSelfContained(path, allowedBasePaths: [allowed]));
    }

    [Fact]
    public void FileReferenceInsideTheAllowlist_ShouldBeAccepted()
    {
        var allowed = Path.Combine(TestDir, "allowed_ok");
        Directory.CreateDirectory(allowed);
        var inside = Path.Combine(allowed, "logo.png");
        File.WriteAllText(inside, "not really a png");

        var uri = "file:///" + inside.Replace("\\", "/");
        var path = WriteMht("file_inside.mht", $"<html><body><img src=\"{uri}\"></body></html>");

        var exception = Record.Exception(() =>
            MhtExternalReferenceScanner.EnsureSelfContained(path, allowedBasePaths: [allowed]));

        Assert.Null(exception);
    }

    [Fact]
    public void UnparsableArchive_ShouldBeRefusedRatherThanWavedThrough()
    {
        // The converter runs its own parser over the same bytes. "We could not read it" is not
        // evidence that nothing in it would be fetched.
        var path = CreateTestFilePath("broken.mht");
        File.WriteAllText(path, "this is not an MHT archive");

        Assert.ThrowsAny<ArgumentException>(() => MhtExternalReferenceScanner.EnsureSelfContained(path));
    }

    [Fact]
    public void SelfContainedArchive_ShouldStillBeAccepted()
    {
        var path = WriteMht("selfcontained.mht",
            "<html><body><img src=\"cid:logo\"><p>hello</p></body></html>");

        var exception = Record.Exception(() =>
            MhtExternalReferenceScanner.EnsureSelfContained(path));

        Assert.Null(exception);
    }
}
