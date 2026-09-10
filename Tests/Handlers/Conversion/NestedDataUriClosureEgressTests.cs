using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R11-SEC01 (§23.3): the decode closure was missing one reading, and that gap fetched.
///     <para>
///         The previous round put every decoded payload back through <c>Readings()</c> — as
///         written, CSS-unescaped, separator-normalised, and both. It left out the one reading the
///         outer part already had: character references. So an inner carrier spelled
///         <c>data&amp;#58;text/html;base64,…</c> was produced by the scanner and never decoded,
///         and the pinned Aspose build fetched what was inside it. Measured with a loopback
///         listener: the scanner did not refuse, and the probe received a real GET.
///     </para>
///     <para>
///         Also here: a media type declaring <c>charset</c> twice. Reading only the first one
///         means the scan decodes by a charset the parser may not use.
///     </para>
/// </summary>
public class NestedDataUriClosureEgressTests : TestBase
{
    /// <summary>Base64 of the given text.</summary>
    /// <param name="text">The text to encode.</param>
    /// <param name="encoding">How to encode it, UTF-8 when omitted.</param>
    /// <returns>The base64 form.</returns>
    private static string Base64(string text, Encoding? encoding = null)
    {
        return System.Convert.ToBase64String((encoding ?? Encoding.UTF8).GetBytes(text));
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
    public void ADataUriSpelledWithACharacterReferenceInsideAPayload_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("nested_charref_data.html");

        // The second level's scheme separator is an entity, so it only becomes a `data:` URI
        // after character references are resolved — the one reading the inner closure omitted.
        var innermost = Base64("<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>");
        var middle = "<iframe src=\"data&#58;text/html;base64," + innermost + "\"></iframe>";

        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;base64," + Base64(middle) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("nested_charref_data.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal,
            "a data: URI spelled with a character reference inside a decoded payload");
    }

    [Fact]
    public void ADataUriRevealedByBothAnEntityAndABackslash_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("nested_charref_backslash.html");

        // Entity-spelled colon at the second level, backslash separators at the third: the
        // closure has to keep applying every reading to whatever the previous one revealed.
        var innermost = Base64("<img src=\"http:" + new string('\\', 2) + "127.0.0.1:"
                               + probe.Port + new string('\\', 1) + "probe.png\">");
        var middle = "<iframe src=\"data&#58;text/html;base64," + innermost + "\"></iframe>";

        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;base64," + Base64(middle) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("nested_charref_backslash.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal,
            "a reference revealed by an entity and then a backslash");
    }

    [Fact]
    public void AMediaTypeDeclaringCharsetTwice_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_double_charset.html");

        // Reading only the first charset means decoding by one the parser may not use. Which one
        // this build honours is not the point: a document whose declaration contradicts itself
        // cannot be decoded with confidence, so it cannot be cleared.
        var inner = "<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>";
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;charset=utf-8;charset=utf-16le;base64,"
                 + Base64(inner, Encoding.Unicode) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_double_charset.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "a media type declaring charset twice");
    }

    [Fact]
    public void AnInertNestedDocument_ShouldStillConvert()
    {
        // The control: widening the closure must not refuse a nested document that references
        // nothing, or the guard has stopped being about egress.
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("nested_inert.html");

        var innermost = Base64("<p>nothing remote</p>");
        var middle = "<iframe src=\"data&#58;text/html;base64," + innermost + "\"></iframe>";

        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;base64," + Base64(middle) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("nested_inert.pdf"));

        Assert.Null(probe.FirstRequest);
        Assert.Null(refusal);
    }
}
