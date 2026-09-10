using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R9-SEC01: the <c>data:</c> document variants the scanner's caps and its pattern do not
///     cover, each decided on a socket rather than by reading the regex.
///     <para>
///         The scanner decodes at most 32 embedded documents and skips any payload over 200,000
///         characters, and both of those report "nothing found" rather than "not fully scanned".
///         Its pattern also accepts only <c>data:&lt;type&gt;;base64,</c>, so a media type carrying
///         a parameter and a percent-encoded payload are not seen at all, and an embedded document
///         is decoded one level deep with no separator normalisation inside it (R9-SEC01).
///     </para>
///     <para>
///         Each case asserts two things: that nothing reached the probe, and that the refusal came
///         from this policy. A conversion that failed for an unrelated reason would otherwise look
///         like the policy holding.
///     </para>
/// </summary>
public class DataDocumentVariantEgressTests : TestBase
{
    /// <summary>How many embedded documents the scanner decodes before its cap.</summary>
    private const int EmbeddedDocumentCap = 32;

    /// <summary>The payload length the scanner refuses to decode past, in characters.</summary>
    private const int PayloadCharCap = 200_000;

    /// <summary>Base64 of the given text.</summary>
    /// <param name="text">The text to encode.</param>
    /// <returns>The base64 form.</returns>
    private static string Base64(string text)
    {
        return System.Convert.ToBase64String(Encoding.UTF8.GetBytes(text));
    }

    /// <summary>Converts one document and reports the refusal, if any.</summary>
    /// <param name="probe">The listening probe.</param>
    /// <param name="input">The document to convert.</param>
    /// <param name="output">Where the PDF would go.</param>
    /// <returns>The refusal message, or null when the conversion was attempted.</returns>
    private static string? Convert(
        EntityEncodedReferenceEgressTests.LoopbackProbe probe, string input, string output)
    {
        string? refusal = null;
        try
        {
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []);
        }
        catch (ArgumentException exception)
        {
            refusal = exception.Message;
        }
        catch (Exception exception)
        {
            refusal = exception.GetType().Name;
        }

        var deadline = DateTime.UtcNow.AddSeconds(2);
        while (probe.FirstRequest == null && DateTime.UtcNow < deadline)
            Thread.Sleep(50);

        return refusal;
    }

    /// <summary>
    ///     Asserts the document was refused by this policy and that nothing reached the probe.
    /// </summary>
    /// <param name="probe">The listening probe.</param>
    /// <param name="refusal">What the conversion reported.</param>
    /// <param name="what">What was being tested, for the failure message.</param>
    private static void ShouldBeRefusedWithoutConnecting(
        EntityEncodedReferenceEgressTests.LoopbackProbe probe, string? refusal, string what)
    {
        Assert.True(probe.FirstRequest == null,
            $"{what} reached the network: {probe.FirstRequest}");
        Assert.NotNull(refusal);
        Assert.Contains("references", refusal, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>An HTML document whose body is the given markup.</summary>
    /// <param name="body">The body markup.</param>
    /// <returns>The document text.</returns>
    private static string Html(string body)
    {
        return "<html><body>" + body + "</body></html>";
    }

    [Fact]
    public void AnEmbeddedDocumentPastTheDecodeCap_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_past_cap.html");

        // The first 32 are inert; the one that carries the reference is the 33rd, which the
        // scanner stops before reaching and reports nothing about.
        var harmless = Base64("<p>nothing here</p>");
        var body = new StringBuilder();
        for (var i = 0; i < EmbeddedDocumentCap; i++)
            body.Append("<iframe src=\"data:text/html;base64,").Append(harmless).Append("\"></iframe>");

        var payload = Base64("<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>");
        body.Append("<iframe src=\"data:text/html;base64,").Append(payload).Append("\"></iframe>");

        File.WriteAllText(input, Html(body.ToString()), Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_past_cap.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "an embedded document past the decode cap");
    }

    [Fact]
    public void AnEmbeddedDocumentPastTheSizeCap_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_past_size.html");

        // Padded past the character cap the scanner will decode, so it skips the payload and
        // reports the document clean.
        var inner = "<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>"
                    + "<!--" + new string('x', PayloadCharCap) + "-->";
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;base64," + Base64(inner) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_past_size.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "an embedded document past the size cap");
    }

    [Fact]
    public void AMediaTypeCarryingAParameter_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_charset.html");

        var payload = Base64("<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>");
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;charset=utf-8;base64," + payload + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_charset.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "a data: URI with a charset parameter");
    }

    [Fact]
    public void APercentEncodedDocument_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_percent.html");

        // Not base64 at all: the payload is percent-encoded, which the pattern does not accept.
        var inner = Uri.EscapeDataString("<img src=\"http://127.0.0.1:" + probe.Port + "/probe.png\">");
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html," + inner + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_percent.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "a percent-encoded data: document");
    }

    [Fact]
    public void ADocumentTwoDataLevelsDeep_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_nested.html");

        var innermost = Base64("<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>");
        var middle = Base64("<iframe src=\"data:text/html;base64," + innermost + "\"></iframe>");
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;base64," + middle + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_nested.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "a document two data: levels deep");
    }

    [Fact]
    public void ABackslashUrlInsideAnEmbeddedDocument_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_inner_backslash.html");

        // Separator normalisation is applied to the outer part but not inside a decoded payload.
        var inner = "<img src=\"http:" + new string('\\', 2) + "127.0.0.1:" + probe.Port
                    + new string('\\', 1) + "probe.png\">";
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;base64," + Base64(inner) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_inner_backslash.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal,
            "a backslash-written URL inside an embedded document");
    }
}
