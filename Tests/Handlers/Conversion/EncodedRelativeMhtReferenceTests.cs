using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R17-M02: whether the pinned renderer resolves a relative reference that arrives encoded.
///     <para>
///         The scanner searches an MHT part's <em>raw</em> text for relative references, while
///         absolute URIs are searched across the decoded readings as well. §28.14 recorded the gap
///         as unproven in both directions: existing tests show a plain relative <c>url()</c> with a
///         remote <c>Content-Location</c> does reach the network, but nothing had established
///         whether the renderer accepts the same reference written quoted-printable or base64 — and
///         if it does not, the scanner has nothing to miss.
///     </para>
///     <para>
///         Asked with a real conversion against a bounded loopback listener. The answer is that the
///         <em>scanner</em> covers these: it decodes the part and finds the reference, and the
///         conversion is refused before anything is fetched. That is worth stating precisely,
///         because the first version of this file asserted only "no request arrived" and swallowed
///         every exception — which is also what a conversion that failed for an unrelated reason
///         looks like, and it was read as evidence about the renderer rather than the scanner
///         (R18-TEST01).
///     </para>
///     <para>
///         So each case now pins the refusal as well as the silence. What is <em>not</em> claimed
///         here is anything about the renderer's own behaviour: nothing reaches it, so this says
///         nothing about what it would do if something did.
///     </para>
/// </summary>
public class EncodedRelativeMhtReferenceTests : TestBase
{
    /// <summary>An MHT whose stylesheet part carries a relative reference in the given encoding.</summary>
    /// <param name="name">The fixture file name.</param>
    /// <param name="port">The loopback port the part's Content-Location points at.</param>
    /// <param name="encoding">The part's Content-Transfer-Encoding.</param>
    /// <param name="body">The part body, already encoded.</param>
    /// <returns>The archive's path.</returns>
    private string AnArchive(string name, int port, string encoding, string body)
    {
        var path = CreateTestFilePath(name);

        File.WriteAllText(path,
            "MIME-Version: 1.0\r\n"
            + "Content-Type: multipart/related; boundary=\"BOUNDARY\"\r\n\r\n"
            + "--BOUNDARY\r\n"
            + "Content-Type: text/html; charset=utf-8\r\n"
            + $"Content-Location: http://127.0.0.1:{port}/index.html\r\n\r\n"
            + "<html><head><link rel=\"stylesheet\" href=\"theme.css\"></head>"
            + "<body><p>a document</p></body></html>\r\n"
            + "--BOUNDARY\r\n"
            + "Content-Type: text/css\r\n"
            + $"Content-Location: http://127.0.0.1:{port}/theme.css\r\n"
            + $"Content-Transfer-Encoding: {encoding}\r\n\r\n"
            + body + "\r\n"
            + "--BOUNDARY--\r\n",
            Encoding.ASCII);

        return path;
    }

    /// <summary>Converts an archive and reports what the probe saw.</summary>
    /// <param name="probe">The listening probe.</param>
    /// <param name="input">The archive.</param>
    /// <param name="output">Where the PDF would go.</param>
    /// <returns>The exception's type and message, or null when the conversion completed.</returns>
    private static (string Type, string Message)? Convert(
        EntityEncodedReferenceEgressTests.LoopbackProbe probe, string input, string output)
    {
        (string, string)? refusal = null;
        try
        {
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []);
        }
        catch (Exception exception)
        {
            refusal = (exception.GetType().Name, exception.Message);
        }

        var deadline = DateTime.UtcNow.AddSeconds(2);
        while (probe.FirstRequest == null && DateTime.UtcNow < deadline)
            Thread.Sleep(50);

        return refusal;
    }

    [Theory]
    [InlineData("quoted-printable", "body { background: url(=22logo.png=22); }")]
    [InlineData("base64", "Ym9keSB7IGJhY2tncm91bmQ6IHVybCgibG9nby5wbmciKTsgfQ==")]
    public void ARelativeReferenceThatArrivesEncoded_ShouldNotReachTheNetwork(
        string encoding, string body)
    {
        // Both spellings of the same stylesheet: `body { background: url("logo.png"); }` written
        // quoted-printable, and written base64. A relative reference the scanner cannot see in the
        // raw text, against a part whose Content-Location makes it resolve to a remote host.
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();

        var input = AnArchive($"encoded_relative_{encoding}.mht", probe.Port, encoding, body);
        var refusal = Convert(probe, input, CreateTestFilePath($"encoded_relative_{encoding}.pdf"));

        Assert.True(probe.FirstRequest == null,
            $"a {encoding} relative reference reached the network: {probe.FirstRequest}");

        // The refusal, not merely the silence. Accepting any failure meant a conversion that broke
        // for an unrelated reason produced the same green (R18-TEST01).
        Assert.NotNull(refusal);
        Assert.Equal("ArgumentException", refusal.Value.Type);
        Assert.Contains("references", refusal.Value.Message, StringComparison.OrdinalIgnoreCase);

        // And the reference it names is the one that was hidden by the encoding, which is what
        // shows the scanner decoded the part rather than tripping over something else.
        Assert.Contains("logo.png", refusal.Value.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APlainRelativeReference_ShouldStillNotReachTheNetwork()
    {
        // The control, and the shape that is known to be reachable: this one the scanner does read
        // out of the raw part text, so it is refused before any request. If this ever stops being
        // refused, the encoded cases above prove nothing.
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();

        var input = AnArchive("plain_relative.mht", probe.Port, "7bit",
            "body { background: url(\"logo.png\"); }");

        var refusal = Convert(probe, input, CreateTestFilePath("plain_relative.pdf"));

        Assert.True(probe.FirstRequest == null,
            $"a plain relative reference reached the network: {probe.FirstRequest}");
        Assert.NotNull(refusal);
        Assert.Equal("ArgumentException", refusal.Value.Type);
        Assert.Contains("logo.png", refusal.Value.Message, StringComparison.Ordinal);
    }
}
