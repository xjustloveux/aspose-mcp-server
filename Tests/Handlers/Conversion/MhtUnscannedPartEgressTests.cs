using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R8-SEC01: the scanner and the converter are two different MIME parsers over the same bytes,
///     so a part one of them drops is a part the other may still act on.
///     <para>
///         Measured: <c>MailMessage</c> exposes a second sibling <c>text/html</c> part as neither
///         an alternate view nor a linked resource — it is simply gone, and with it any reference
///         it carried. Whether that matters depends on what the converter's own parser does, which
///         only a real conversion against a listening socket can answer.
///     </para>
/// </summary>
public class MhtUnscannedPartEgressTests : TestBase
{
    [Fact]
    public void AReferenceInAPartTheScannerCannotSee_ShouldNotReachTheNetwork()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();

        var path = CreateTestFilePath("unscanned_part.mht");
        File.WriteAllText(path, string.Join("\r\n",
            "From: <fixture>", "Subject: fixture", "MIME-Version: 1.0",
            "Content-Type: multipart/related; boundary=\"B\"", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://127.0.0.1:" + probe.Port + "/index.html", "",
            "<html><body>first</body></html>", "",
            "--B", "Content-Type: text/html; charset=\"utf-8\"",
            "Content-Transfer-Encoding: 8bit",
            "Content-Location: http://127.0.0.1:" + probe.Port + "/page.html", "",
            "<html><body><img src=\"http://127.0.0.1:" + probe.Port + "/probe.png\"></body></html>", "",
            "--B--", ""), new UTF8Encoding(false));

        // The scanner runs first, as it does in the conversion path.
        var refusal = Record.Exception(() => MhtExternalReferenceScanner.EnsureSelfContained(path));

        if (refusal == null)
            try
            {
                DocumentConverter.ConvertToPdfFromSpecialFormat(
                    path, CreateTestFilePath("unscanned_part.pdf"), []);
            }
            catch (Exception)
            {
                // A conversion failure is not the measurement; the socket is.
            }

        Thread.Sleep(500);

        Assert.True(probe.FirstRequest == null,
            "A part the scanner never saw produced an outbound request: " + probe.FirstRequest);
    }
}
