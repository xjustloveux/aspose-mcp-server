using System.IO.Compression;
using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     Every convertible input must be refused before it can make the server fetch something
///     (§13.7 item 2).
///     <para>
///         Measured with a localhost probe server against the pinned Aspose.Pdf: converting HTML,
///         Markdown, SVG and EPUB each fetched the URL the document named. For HTML the converter
///         did pass <c>CustomLoaderOfExternalResources</c>, and an instrumented strategy recorded
///         <em>zero</em> invocations while the probe still logged the request — the callback is not
///         consulted for an <c>img src</c> on that path, so it protected nothing. LaTeX did not
///         fetch. The control that works is the one MHT already used: refuse the input before it is
///         opened.
///     </para>
///     <para>
///         These tests need no network. Refusal happens before the converter opens the file, so a
///         reachable host is never required; the URL only has to be present in the document.
///     </para>
/// </summary>
public class ConvertibleFormatEgressTests : TestBase
{
    private const string RemoteUrl = "http://127.0.0.1:9/probe.png";

    /// <summary>
    ///     Writes a self-contained EPUB, optionally carrying a remote image reference.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <param name="body">Body markup for the single page.</param>
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
            "<?xml version=\"1.0\"?><container version=\"1.0\" " +
            "xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\"><rootfiles>" +
            "<rootfile full-path=\"content.opf\" media-type=\"application/oebps-package+xml\"/>" +
            "</rootfiles></container>");
        Add("content.opf",
            "<?xml version=\"1.0\"?><package xmlns=\"http://www.idpf.org/2007/opf\" version=\"2.0\" " +
            "unique-identifier=\"i\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\">" +
            "<dc:title>t</dc:title><dc:identifier id=\"i\">i</dc:identifier><dc:language>en</dc:language>" +
            "</metadata><manifest><item id=\"p\" href=\"page.xhtml\" " +
            "media-type=\"application/xhtml+xml\"/></manifest><spine><itemref idref=\"p\"/></spine></package>");
        Add("page.xhtml",
            "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><body>" + body +
            "</body></html>");

        return path;
    }

    /// <summary>
    ///     A document naming a remote resource must not reach the converter.
    /// </summary>
    /// <param name="fileName">Fixture file name, which selects the format.</param>
    /// <param name="content">Document content carrying the remote reference.</param>
    [Theory]
    [InlineData("egress.html", "<html><body><img src=\"" + RemoteUrl + "\"></body></html>")]
    [InlineData("egress.md", "# t\n\n![x](" + RemoteUrl + ")\n")]
    [InlineData("egress.svg",
        "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\">" +
        "<image xlink:href=\"" + RemoteUrl + "\" width=\"10\" height=\"10\"/></svg>")]
    public void ConvertingADocumentThatNamesARemoteResource_ShouldBeRefused(string fileName, string content)
    {
        var input = CreateTestFilePath(fileName);
        File.WriteAllText(input, content, Encoding.UTF8);
        var output = CreateTestFilePath(fileName + ".pdf");

        var exception = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []));

        Assert.Contains("references resources", exception.Message);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void ConvertingAnEpubThatNamesARemoteResource_ShouldBeRefused()
    {
        var input = WriteEpub(CreateTestFilePath("egress.epub"), "<img src=\"" + RemoteUrl + "\"/>");
        var output = CreateTestFilePath("egress_epub.pdf");

        Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []));
        Assert.False(File.Exists(output));
    }

    /// <summary>
    ///     The refusal must not extend to documents that reference nothing, or the guard would
    ///     simply disable the feature.
    /// </summary>
    /// <param name="fileName">Fixture file name, which selects the format.</param>
    /// <param name="content">Self-contained document content.</param>
    [Theory]
    [InlineData("clean.html", "<html><body><h1>Self contained</h1></body></html>")]
    [InlineData("clean.md", "# Title\n\nPlain paragraph.\n")]
    [InlineData("clean.svg",
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"40\" height=\"40\">" +
        "<rect width=\"30\" height=\"30\" fill=\"#333\"/></svg>")]
    public void ConvertingASelfContainedDocument_ShouldStillWork(string fileName, string content)
    {
        var input = CreateTestFilePath(fileName);
        File.WriteAllText(input, content, Encoding.UTF8);
        var output = CreateTestFilePath(fileName + ".pdf");

        DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []);

        Assert.True(File.Exists(output));
        Assert.True(new FileInfo(output).Length > 0);
    }

    [Fact]
    public void OptingIn_ShouldAllowTheRemoteReferenceThrough()
    {
        // allowExternalResources is a deployment-layer decision; when it is set the scan is skipped
        // and the converter is allowed to do whatever it does with the reference.
        var input = CreateTestFilePath("optin.md");
        File.WriteAllText(input, "# t\n\n![x](" + RemoteUrl + ")\n", Encoding.UTF8);
        var output = CreateTestFilePath("optin.pdf");

        // Port 9 is the discard port: nothing listens, so this exercises the opt-in path without
        // making a request that could succeed.
        var recorded = Record.Exception(() =>
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, [], true));

        Assert.True(recorded is null or not ArgumentException,
            $"Opting in must not be refused by the scanner, but got: {recorded}");
    }
}
