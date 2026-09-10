using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R10-SEC01 (§22.4): the <c>data:</c> decoder was not a closure, and it ignored charset.
///     <para>
///         Every base64 payload was decoded as UTF-8 whatever the media type declared, so a
///         document written as <c>charset=utf-16le</c> came back as UTF-8 nonsense carrying NUL
///         bytes: the scanner produced a string, found no URL in it, and reported the archive
///         clean. And while the <em>outer</em> part was searched for <c>data:</c> URIs in all four
///         of its readings, a decoded payload was only searched as written — so a second level
///         revealed by the payload's own CSS escapes or backslashes was produced and then never
///         decoded.
///     </para>
///     <para>
///         Both are scanner-proof gaps: whether the pinned Aspose build would itself dereference
///         these shapes is decided on a socket, and each test asserts both outcomes — nothing
///         reached the probe, and the refusal came from this policy.
///     </para>
/// </summary>
public class DataUriCharsetClosureTests : TestBase
{
    /// <summary>Base64 of the given text in the given encoding.</summary>
    /// <param name="text">The text to encode.</param>
    /// <param name="encoding">How to encode it.</param>
    /// <returns>The base64 form.</returns>
    private static string Base64(string text, Encoding encoding)
    {
        return System.Convert.ToBase64String(encoding.GetBytes(text));
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
    public void AUtf16EmbeddedDocument_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_utf16.html");

        var inner = "<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>";
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;charset=utf-16le;base64,"
                 + Base64(inner, Encoding.Unicode) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_utf16.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "a UTF-16 embedded document");
    }

    [Fact]
    public void AnUnknownCharset_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_unknown_charset.html");

        var inner = "<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>";
        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;charset=windows-1252;base64,"
                 + Base64(inner, Encoding.UTF8) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_unknown_charset.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal, "an embedded document in an unknown charset");
    }

    [Fact]
    public void ANestedDataUriRevealedOnlyInsideTheInnerDocument_ShouldBeRefusedWithoutConnecting()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_inner_revealed.html");

        // The second level is spelled with a CSS escape *inside* the first level's payload, so it
        // exists only after that payload is unescaped — which the decoder never did.
        var innermost = Base64("<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>",
            Encoding.UTF8);
        var middle = "<iframe src=\"dat\\61 :text/html;base64," + innermost + "\"></iframe>";

        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;base64," + Base64(middle, Encoding.UTF8)
                                                        + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_inner_revealed.pdf"));
        ShouldBeRefusedWithoutConnecting(probe, refusal,
            "a nested data: URI revealed only inside the inner document");
    }

    [Fact]
    public void ACharsetDeclaredWithoutNamingOne_ShouldBeRefused()
    {
        // R13-SEC02. A media type that says `charset=` has made a declaration this scan cannot
        // honour, but it read back as the same empty string as "no charset was declared" — and
        // that is the one answer that skips the charset gate and decodes as UTF-8 regardless. So
        // the document that most explicitly declined to say how it should be read was the one
        // read with the fewest questions asked, and the scan and the parser could be looking at
        // different text.
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_empty_charset.html");

        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;charset=;base64,"
                 + Base64("<p>nothing remote</p>", Encoding.UTF8) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_empty_charset.pdf"));

        Assert.True(probe.FirstRequest == null,
            "an embedded document declaring no character set reached the network: "
            + probe.FirstRequest);
        Assert.NotNull(refusal);
        Assert.Contains("without naming one", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void ACharsetDeclaredAsAnEmptyQuotedString_ShouldBeRefused()
    {
        // The same declaration written the other legal way. Quoting it does not make it a name,
        // and the quotes are stripped before the gate sees the value.
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_quoted_empty_charset.html");

        File.WriteAllText(input,
            Html("<iframe src='data:text/html;charset=\"\";base64,"
                 + Base64("<p>nothing remote</p>", Encoding.UTF8) + "'></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_quoted_empty_charset.pdf"));

        Assert.True(probe.FirstRequest == null,
            "an embedded document declaring no character set reached the network: "
            + probe.FirstRequest);
        Assert.NotNull(refusal);
        Assert.Contains("without naming one", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void AnExplicitUtf8Charset_ShouldStillConvert()
    {
        // The control: refusing an unreadable charset must not refuse the ordinary one that says
        // what it is.
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_utf8_charset.html");

        File.WriteAllText(input,
            Html("<iframe src=\"data:text/html;charset=utf-8;base64,"
                 + Base64("<p>nothing remote</p>", Encoding.UTF8) + "\"></iframe>"),
            Encoding.UTF8);

        var refusal = Convert(probe, input, CreateTestFilePath("data_utf8_charset.pdf"));

        Assert.Null(probe.FirstRequest);
        Assert.Null(refusal);
    }
}
