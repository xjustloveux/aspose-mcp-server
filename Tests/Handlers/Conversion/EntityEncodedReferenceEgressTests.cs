using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     R7-SEC01: the external-reference scanner matches literal URIs in raw text, but HTML, XHTML
///     and SVG are parsed after that, and a parser decodes numeric character references. A
///     reference written as <c>http&amp;#x3a;&amp;#x2f;&amp;#x2f;host/x.png</c> carries no literal
///     <c>://</c> for the scanner to find, and the question is whether anything downstream turns it
///     back into a URL and fetches it.
///     <para>
///         The claim being tested is about a side effect, so the test measures the side effect: a
///         loopback socket is opened, the document names it through an encoded reference, and the
///         conversion is attempted. A connection arriving at that socket is egress no matter what
///         the scanner reported, and no connection is evidence that this shape does not reach the
///         network on the pinned Aspose build. Asserting only that an exception was thrown would
///         have proved neither.
///     </para>
/// </summary>
public class EntityEncodedReferenceEgressTests : TestBase
{
    /// <summary>
    ///     Writes the probe URL with its scheme punctuation as numeric character references, so no
    ///     literal <c>://</c> appears in the bytes the scanner reads.
    /// </summary>
    /// <param name="port">Port the probe is listening on.</param>
    /// <returns>The encoded URL.</returns>
    private static string EncodedUrl(int port)
    {
        return "http&#x3a;&#x2f;&#x2f;127.0.0.1:" + port + "&#x2f;probe.png";
    }

    /// <summary>Writes a minimal EPUB whose single page carries the given body markup.</summary>
    /// <param name="path">Destination path.</param>
    /// <param name="body">Body markup for the page.</param>
    /// <returns>The path written.</returns>
    private static string WriteEpub(string path, string body)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);

        void Add(string name, string content)
        {
            using var writer = new StreamWriter(archive.CreateEntry(name).Open(), Encoding.UTF8);
            writer.Write(content);
        }

        Add("mimetype", "application/epub+zip");
        Add("META-INF/container.xml",
            "<?xml version=\"1.0\"?><container version=\"1.0\" "
            + "xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>"
            + "<rootfile full-path=\"content.opf\" media-type=\"application/oebps-package+xml\"/>"
            + "</rootfiles></container>");
        Add("content.opf",
            "<?xml version=\"1.0\"?><package xmlns=\"http://www.idpf.org/2007/opf\" version=\"2.0\" "
            + "unique-identifier=\"i\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">"
            + "<dc:title>t</dc:title><dc:identifier id=\"i\">i</dc:identifier><dc:language>en</dc:language>"
            + "</metadata><manifest><item id=\"p\" href=\"page.xhtml\" "
            + "media-type=\"application/xhtml+xml\"/></manifest><spine><itemref idref=\"p\"/></spine></package>");
        Add("page.xhtml",
            "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><body>" + body
            + "</body></html>");

        return path;
    }

    /// <summary>Writes a minimal MHT archive whose HTML part carries the given body markup.</summary>
    /// <param name="path">Destination path.</param>
    /// <param name="body">Body markup for the page.</param>
    /// <returns>The path written.</returns>
    private static string WriteMht(string path, string body)
    {
        var content = "From: <probe@localhost>\r\n"
                      + "Subject: probe\r\n"
                      + "MIME-Version: 1.0\r\n"
                      + "Content-Type: multipart/related; boundary=\"BOUND\"\r\n"
                      + "\r\n"
                      + "--BOUND\r\n"
                      + "Content-Type: text/html; charset=\"utf-8\"\r\n"
                      + "Content-Transfer-Encoding: 8bit\r\n"
                      + "Content-Location: file:///probe/index.html\r\n"
                      + "\r\n"
                      + "<html><body>" + body + "</body></html>\r\n"
                      + "--BOUND--\r\n";

        File.WriteAllText(path, content, new UTF8Encoding(false));
        return path;
    }

    /// <summary>
    ///     Attempts the conversion, ignoring a refusal, and reports whether the probe was contacted.
    /// </summary>
    /// <param name="probe">The listening probe.</param>
    /// <param name="input">Input document path.</param>
    /// <param name="output">Output PDF path.</param>
    /// <returns>The refusal message, or null when the conversion was attempted.</returns>
    private static string? ConvertAndSettle(LoopbackProbe probe, string input, string output)
    {
        string? refusal = null;
        try
        {
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []);
        }
        catch (ArgumentException ex)
        {
            refusal = ex.Message;
        }
        catch (Exception ex)
        {
            // A malformed-input failure is not the measurement either way; the probe is.
            refusal = ex.GetType().Name + ": " + ex.Message;
        }

        // A fetch issued during conversion has already happened by the time it returns, but give a
        // late connection a bounded moment to arrive rather than racing it.
        var deadline = DateTime.UtcNow.AddSeconds(2);
        while (probe.FirstRequest == null && DateTime.UtcNow < deadline)
            Thread.Sleep(50);

        return refusal;
    }

    [Fact]
    public void HtmlWithAnEntityEncodedRemoteReference_ShouldNotReachTheNetwork()
    {
        using var probe = new LoopbackProbe();
        var input = CreateTestFilePath("entity_encoded.html");
        File.WriteAllText(input,
            "<html><body><img src=\"" + EncodedUrl(probe.Port) + "\"></body></html>", Encoding.UTF8);

        var refusal = ConvertAndSettle(probe, input, CreateTestFilePath("entity_encoded_html.pdf"));

        Assert.True(probe.FirstRequest == null,
            "An entity-encoded reference reached the network: " + probe.FirstRequest);

        // And the refusal has to be this policy's. Asserting only that nothing connected would
        // pass for a conversion that failed to parse the document at all (R8-T01).
        Assert.NotNull(refusal);
        Assert.Contains("references resources", refusal);
    }

    [Fact]
    public void SvgWithAnEntityEncodedRemoteReference_ShouldNotReachTheNetwork()
    {
        using var probe = new LoopbackProbe();
        var input = CreateTestFilePath("entity_encoded.svg");
        File.WriteAllText(input,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\">"
            + "<image xlink:href=\"" + EncodedUrl(probe.Port) + "\" width=\"10\" height=\"10\"/></svg>",
            Encoding.UTF8);

        var refusal = ConvertAndSettle(probe, input, CreateTestFilePath("entity_encoded_svg.pdf"));

        Assert.True(probe.FirstRequest == null,
            "An entity-encoded reference reached the network: " + probe.FirstRequest);

        // And the refusal has to be this policy's. Asserting only that nothing connected would
        // pass for a conversion that failed to parse the document at all (R8-T01).
        Assert.NotNull(refusal);
        Assert.Contains("references resources", refusal);
    }

    [Fact]
    public void EpubWithAnEntityEncodedRemoteReference_ShouldNotReachTheNetwork()
    {
        using var probe = new LoopbackProbe();
        var input = WriteEpub(CreateTestFilePath("entity_encoded.epub"),
            "<img src=\"" + EncodedUrl(probe.Port) + "\"/>");

        var refusal = ConvertAndSettle(probe, input, CreateTestFilePath("entity_encoded_epub.pdf"));

        Assert.True(probe.FirstRequest == null,
            "An entity-encoded reference reached the network: " + probe.FirstRequest);

        // And the refusal has to be this policy's. Asserting only that nothing connected would
        // pass for a conversion that failed to parse the document at all (R8-T01).
        Assert.NotNull(refusal);
        Assert.Contains("references resources", refusal);
    }

    [Fact]
    public void MhtWithAnEntityEncodedRemoteReference_ShouldNotReachTheNetwork()
    {
        using var probe = new LoopbackProbe();
        var input = WriteMht(CreateTestFilePath("entity_encoded.mht"),
            "<img src=\"" + EncodedUrl(probe.Port) + "\">");

        ConvertAndSettle(probe, input, CreateTestFilePath("entity_encoded_mht.pdf"));

        Assert.True(probe.FirstRequest == null,
            "An entity-encoded reference reached the network: " + probe.FirstRequest);
    }

    [Fact]
    public void ThePlainSpellingOfTheSameReference_ShouldStillBeRefused()
    {
        // The control: with the same probe and the same document shape, the literal form is what
        // the scanner already catches. If this ever stops refusing, the encoded cases above are
        // measuring nothing.
        using var probe = new LoopbackProbe();
        var input = CreateTestFilePath("plain_reference.html");
        File.WriteAllText(input,
            "<html><body><img src=\"http://127.0.0.1:" + probe.Port + "/probe.png\"></body></html>",
            Encoding.UTF8);

        var refusal = ConvertAndSettle(probe, input, CreateTestFilePath("plain_reference.pdf"));

        Assert.NotNull(refusal);
        Assert.Contains("references resources", refusal);
        Assert.Null(probe.FirstRequest);
    }

    /// <summary>
    ///     A loopback socket that records whether anything connected to it, used as the observable
    ///     end of an outbound request.
    /// </summary>
    internal sealed class LoopbackProbe : IDisposable
    {
        private readonly Task _accepting;
        private readonly CancellationTokenSource _cancellation = new();
        private readonly TcpListener _listener;
        private volatile string? _firstRequest;

        /// <summary>Starts listening on an ephemeral loopback port.</summary>
        internal LoopbackProbe()
        {
            _listener = new TcpListener(IPAddress.Loopback, 0);
            _listener.Start();
            Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
            _accepting = Task.Run(AcceptAsync);
        }

        /// <summary>The port the document should be pointed at.</summary>
        internal int Port { get; }

        /// <summary>The first request line seen, or null when nothing connected.</summary>
        internal string? FirstRequest => _firstRequest;

        /// <inheritdoc />
        public void Dispose()
        {
            _cancellation.Cancel();
            _listener.Stop();

            try
            {
                _accepting.Wait(TimeSpan.FromSeconds(5));
            }
            catch (AggregateException)
            {
                /* best-effort */
            }

            _cancellation.Dispose();
        }

        private async Task AcceptAsync()
        {
            try
            {
                while (!_cancellation.IsCancellationRequested)
                {
                    using var client = await _listener.AcceptTcpClientAsync(_cancellation.Token);
                    await using var stream = client.GetStream();

                    var buffer = new byte[512];
                    var read = await stream.ReadAsync(buffer, _cancellation.Token);
                    Interlocked.CompareExchange(ref _firstRequest,
                        Encoding.ASCII.GetString(buffer, 0, Math.Max(read, 0)), null);

                    // Answer something valid enough that the caller does not sit waiting; the
                    // connection itself is the measurement.
                    var body = "not found"u8.ToArray();
                    var head = Encoding.ASCII.GetBytes(
                        "HTTP/1.1 404 Not Found\r\nContent-Length: " + body.Length
                                                                     + "\r\nConnection: close\r\n\r\n");
                    await stream.WriteAsync(head, _cancellation.Token);
                    await stream.WriteAsync(body, _cancellation.Token);
                }
            }
            catch (OperationCanceledException)
            {
                /* shutting down */
            }
            catch (SocketException)
            {
                /* listener closed */
            }
            catch (ObjectDisposedException)
            {
                /* listener closed */
            }
        }
    }
}
