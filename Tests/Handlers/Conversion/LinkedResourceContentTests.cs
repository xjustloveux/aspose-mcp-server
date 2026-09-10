using System.Text;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R9-SEC02: a linked resource is a part with content, not merely a name.
///     <para>
///         The scan read each linked resource's <c>Content-Location</c> so the archive could be
///         credited with carrying that address, and then looked no further. A stylesheet's
///         <c>@import</c> and <c>url()</c>, an SVG's <c>href</c> and an HTML fragment's <c>src</c>
///         are references the converter will follow, and none of them was ever read.
///     </para>
///     <para>
///         Also here: a part the parser hands back with no stream at all. That returned an empty
///         string, so the archive was scanned without it and reported clean — the one outcome this
///         guard may never produce.
///     </para>
/// </summary>
[Collection("SerialStaticSeams")]
public class LinkedResourceContentTests : TestBase
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
    public void AStylesheetLinkedResource_ShouldBeScannedForItsOwnReferences()
    {
        var path = Archive("linked_css.mht",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/index.html", "",
            "<html><head><link rel=\"stylesheet\" href=\"style.css\"></head><body>x</body></html>", "",
            "--B", "Content-Type: text/css; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/style.css", "",
            "body{background:url(http://tracker.example/pixel.png)}", "",
            "--B--");

        var refusal = Refusal(path);

        Assert.NotNull(refusal);
        Assert.Contains("references resources", refusal, StringComparison.Ordinal);
        Assert.Contains("pixel.png", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void AnSvgLinkedResource_ShouldBeScannedForItsOwnReferences()
    {
        var path = Archive("linked_svg.mht",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/index.html", "",
            "<html><body><img src=\"art.svg\"></body></html>", "",
            "--B", "Content-Type: image/svg+xml; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/art.svg", "",
            "<svg xmlns=\"http://www.w3.org/2000/svg\">",
            "<image href=\"http://tracker.example/pixel.png\"/></svg>", "",
            "--B--");

        var refusal = Refusal(path);

        Assert.NotNull(refusal);
        Assert.Contains("references resources", refusal, StringComparison.Ordinal);
        Assert.Contains("pixel.png", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void ABinaryImageLinkedResource_ShouldNotBeRefused()
    {
        // The control. An image carries no references a converter re-parses, so scanning it would
        // spend the decoding budget to find nothing — and refusing it would break every ordinary
        // saved page.
        var png = Convert.ToBase64String(
            [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x01, 0x02, 0x03]);

        var path = Archive("linked_png.mht",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://localhost/index.html", "",
            "<html><body><img src=\"logo.png\"></body></html>", "",
            "--B", "Content-Type: image/png",
            "Content-Transfer-Encoding: base64",
            "Content-Location: http://localhost/logo.png", "",
            png, "",
            "--B--");

        Assert.Null(Refusal(path));
    }

    [Fact]
    public void APartWithNoStreamAtAll_ShouldBeRefused()
    {
        // The parser hands back its own stream and nothing in a crafted archive makes it null, so
        // the seam is the only way to reach the path — and the path is the whole point of the rule.
        var path = Archive("no_stream.mht",
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
            MhtExternalReferenceScanner.ReadStreamOf = _ => null;

            var refusal = Refusal(path);

            Assert.NotNull(refusal);
            Assert.Contains("no readable content", refusal, StringComparison.Ordinal);
        }
        finally
        {
            MhtExternalReferenceScanner.ReadStreamOf = original;
        }
    }
}
