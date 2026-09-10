using System.IO.Compression;
using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R8-SEC02: the shapes whose canonicalisation the scanner and the parsers might disagree
///     about, each decided by what actually happens on a socket rather than by reading the regex.
///     <para>
///         A CSS escape, a scheme written with backslashes, an active <c>data:</c> document and a
///         ZIP container the reader has no special case for are all ways a reference can exist in
///         a document without appearing to this scanner as a literal URI. Whether any of them
///         reaches the network is a property of the pinned Aspose parsers, so each one is put in
///         front of a listening socket and converted.
///     </para>
///     <para>
///         A case that does not connect is recorded as such, with the version it was measured on;
///         it is not evidence that the shape is safe in general, only that this build does not
///         dereference it.
///     </para>
/// </summary>
public class ParserCanonicalisationEgressTests : TestBase
{
    /// <summary>Converts one document and reports whether the probe was contacted.</summary>
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

    [Fact]
    public void ACssEscapedScheme_ShouldNotReachTheNetwork()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("css_escape.html");

        // CSS unescapes \3a into a colon, so the scheme is not spelled out in the bytes.
        File.WriteAllText(input,
            "<html><head><style>body{background:url(http\\3a //127.0.0.1:" + probe.Port
                                                                           + "/probe.png)}</style></head><body>x</body></html>",
            Encoding.UTF8);

        Convert(probe, input, CreateTestFilePath("css_escape.pdf"));

        Assert.True(probe.FirstRequest == null,
            "A CSS-escaped scheme reached the network: " + probe.FirstRequest);
    }

    [Fact]
    public void ASchemeWrittenWithBackslashes_ShouldNotReachTheNetwork()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("backslash_scheme.html");

        File.WriteAllText(input,
            "<html><body><img src=\"http:\\\\127.0.0.1:" + probe.Port + "\\probe.png\"></body></html>",
            Encoding.UTF8);

        Convert(probe, input, CreateTestFilePath("backslash_scheme.pdf"));

        Assert.True(probe.FirstRequest == null,
            "A backslash-written scheme reached the network: " + probe.FirstRequest);
    }

    [Fact]
    public void AnActiveDataDocument_ShouldNotReachTheNetwork()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_document.html");

        // The data: URI is local by definition, but the document inside it is not: it carries its
        // own remote reference, which the scanner's data: exemption would wave through.
        var inner = "<img src='http://127.0.0.1:" + probe.Port + "/probe.png'>";
        var encoded = System.Convert.ToBase64String(Encoding.UTF8.GetBytes(inner));
        File.WriteAllText(input,
            "<html><body><iframe src=\"data:text/html;base64," + encoded + "\"></iframe></body></html>",
            Encoding.UTF8);

        Convert(probe, input, CreateTestFilePath("data_document.pdf"));

        Assert.True(probe.FirstRequest == null,
            "A document embedded in a data: URI reached the network: " + probe.FirstRequest);
    }

    [Fact]
    public void AnSvgInsideADataUri_ShouldNotReachTheNetwork()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("data_svg.html");

        var inner = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>"
                    + "<image xlink:href='http://127.0.0.1:" + probe.Port + "/probe.png' "
                    + "width='10' height='10'/></svg>";
        var encoded = System.Convert.ToBase64String(Encoding.UTF8.GetBytes(inner));
        File.WriteAllText(input,
            "<html><body><img src=\"data:image/svg+xml;base64," + encoded + "\"></body></html>",
            Encoding.UTF8);

        Convert(probe, input, CreateTestFilePath("data_svg.pdf"));

        Assert.True(probe.FirstRequest == null,
            "An SVG embedded in a data: URI reached the network: " + probe.FirstRequest);
    }

    [Fact]
    public void AZipContainerTheReaderHasNoCaseFor_ShouldNotReachTheNetwork()
    {
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("container.xps");

        // ReadConvertibleParts special-cases EPUB and nothing else, so an XPS is read as raw
        // bytes: a reference inside its compressed entries is not text the scan can see.
        using (var archive = ZipFile.Open(input, ZipArchiveMode.Create))
        {
            using var writer = new StreamWriter(archive.CreateEntry("Documents/1/Pages/1.fpage").Open(),
                Encoding.UTF8);
            writer.Write("<FixedPage xmlns=\"http://schemas.microsoft.com/xps/2005/06\" "
                         + "Width=\"100\" Height=\"100\"><Path><Path.Fill><ImageBrush "
                         + "ImageSource=\"http://127.0.0.1:" + probe.Port + "/probe.png\"/>"
                         + "</Path.Fill></Path></FixedPage>");
        }

        var refusal = Convert(probe, input, CreateTestFilePath("container.pdf"));

        Assert.True(probe.FirstRequest == null,
            "A reference inside a compressed container reached the network: " + probe.FirstRequest);

        // Measured: the converter cannot load this container at all (FileLoadException), and the
        // reference uses XPS's own ImageSource attribute, which this scan's attribute vocabulary
        // does not include. So there is no egress to block here and no refusal to claim — the
        // container is read entry by entry now, but that is shown by the fixture below, not by
        // this one.
        Assert.Equal("FileLoadException", refusal);
    }

    [Fact]
    public void AZipContainerCarryingAnOrdinaryReference_ShouldBeRefused()
    {
        // The change the case above cannot show: every ZIP is read entry by entry now, so a
        // reference inside a compressed entry of a container that is not an EPUB is text the scan
        // can see. Before, only .epub was unpacked and everything else was read as one blob of
        // compressed bytes in which no URL is visible (R8-SEC02).
        using var probe = new EntityEncodedReferenceEgressTests.LoopbackProbe();
        var input = CreateTestFilePath("packaged.xps");

        using (var archive = ZipFile.Open(input, ZipArchiveMode.Create))
        {
            using var writer = new StreamWriter(archive.CreateEntry("content/page.html").Open(),
                Encoding.UTF8);
            writer.Write("<html><body><img src=\"http://127.0.0.1:" + probe.Port
                                                                    + "/probe.png\"></body></html>");
        }

        var refusal = Convert(probe, input, CreateTestFilePath("packaged.pdf"));

        Assert.NotNull(refusal);
        Assert.Contains("references resources", refusal);
        Assert.Null(probe.FirstRequest);
    }
}
