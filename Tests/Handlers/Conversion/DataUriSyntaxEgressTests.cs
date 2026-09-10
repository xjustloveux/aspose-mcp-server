using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R9-SEC01 (§21.4): three more spellings the <c>data:</c> decoder does not read the way a
///     parser does, each decided on a socket rather than by reading the regex.
///     <para>
///         The decoder accepts three media types, but a linked resource may be <c>text/css</c> —
///         so a stylesheet carried in a <c>data:</c> URI is never decoded. It decides "this is
///         base64" by looking for the substring anywhere in the parameters, so a parameter merely
///         *named* something ending in <c>base64</c> sends a percent-encoded payload down the
///         base64 path, where the decode fails and the payload is treated as empty. And it looks
///         for <c>data:</c> URIs only in the raw text, not in the forms that appear after CSS
///         escapes are resolved or backslashes normalised.
///     </para>
///     <para>
///         Each case asserts both that nothing reached the probe and that the refusal came from
///         this policy, so a conversion that failed for an unrelated reason cannot pass as the
///         guard holding.
///     </para>
/// </summary>
public class DataUriSyntaxEgressTests : TestBase
{
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

    /// <summary>Asserts the document was refused by this policy without connecting.</summary>
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
    public void AStylesheetCarriedInADataUri_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_css.html");

        // text/css is an active type when it arrives as a linked resource, but the data: decoder
        // did not accept it, so the same stylesheet inside a data: URI was never decoded.
        var css = "@import url(http://127.0.0.1:" + probe.Port + "/probe.css);";
        File.WriteAllText(input,
            "<html><head><link rel=\"stylesheet\" href=\"data:text/css;base64," + Base64(css)
            + "\"></head><body>x</body></html>", Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_css.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "a stylesheet carried in a data: URI");
    }

    [Fact]
    public void AParameterMerelyNamedLikeBase64_ShouldNotSendAPercentEncodedPayloadDownThatPath()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_pseudo_base64.html");

        // `Contains("base64")` matches this parameter, so a percent-encoded payload was decoded as
        // base64, the decode failed, and the failure was reported as "no content" rather than as
        // "could not be checked".
        var inner = Uri.EscapeDataString("<img src=\"http://127.0.0.1:" + probe.Port + "/probe.png\">");
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;charset=base64," + inner + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_pseudo_base64.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal,
            "a data: URI whose parameter is merely named like base64");
    }

    [Fact]
    public void ADataUriRevealedOnlyByCssUnescaping_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_after_css_escape.html");

        // The embedded-document search ran on the raw text only, so a data: URI that exists only
        // after CSS escapes are resolved was never looked inside.
        var payload = Base64("<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>");
        File.WriteAllText(input,
            Html("<iframe src=\"dat\\61 :text/html;base64," + payload + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_after_css_escape.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal,
            "a data: URI revealed only by CSS unescaping");
    }

    [Fact]
    public void AnOrdinaryInertDataDocument_ShouldStillConvert()
    {
        // The control. These guards refuse "could not be checked"; a data: document that carries
        // no remote reference must still convert, or the refusal is not about egress at all.
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_inert.html");

        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;base64," + Base64("<p>nothing remote</p>")
                                                        + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_inert.pdf"));

        Assert.Null(probe.FirstRequest);
        Assert.Null(refusal);
    }
}
