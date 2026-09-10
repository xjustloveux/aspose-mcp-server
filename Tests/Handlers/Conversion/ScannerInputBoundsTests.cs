using System.IO.Compression;
using System.Text;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Conversion;

/// <summary>
///     The external-reference scan has to decide before the work is done, and it has to decide for
///     the whole document.
///     <para>
///         Two ways it did neither: the MHT limits were applied to what the MIME parser produced,
///         after that parser had read and decoded the entire archive (R4-S02); and a container
///         entry that failed to decode returned the parts collected so far, so a document whose
///         first entries were harmless and whose later one was corrupt scanned clean and went on
///         to the converter, which runs a different parser and may get further (R4-S03).
///     </para>
/// </summary>
public class ScannerInputBoundsTests : TestBase
{
    /// <summary>
    ///     Writes an EPUB whose entries are added one by one, so a deliberately corrupt entry can
    ///     be placed after a valid one.
    /// </summary>
    /// <param name="path">Destination path.</param>
    /// <param name="corruptTail">Whether to append an entry whose deflate stream is invalid.</param>
    /// <returns>The path written.</returns>
    private static string WriteEpub(string path, bool corruptTail)
    {
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
        {
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
                "<dc:title>t</dc:title><dc:identifier id=\"i\">i</dc:identifier>" +
                "<dc:language>en</dc:language></metadata><manifest>" +
                "<item id=\"p\" href=\"page.xhtml\" media-type=\"application/xhtml+xml\"/></manifest>" +
                "<spine><itemref idref=\"p\"/></spine></package>");
            Add("page.xhtml",
                "<?xml version=\"1.0\"?><html xmlns=\"http://www.w3.org/1999/xhtml\"><body>" +
                "<p>nothing remote here</p></body></html>");

            if (corruptTail)
                Add("tail.xhtml", "<html><body><p>placeholder</p></body></html>");
        }

        if (!corruptTail) return path;

        // Corrupt the last entry's compressed bytes in place. The central directory still lists
        // it, so opening the archive succeeds and only reading that entry fails — which is
        // exactly the shape that produced a clean partial scan.
        var bytes = File.ReadAllBytes(path);
        var marker = "tail.xhtml"u8.ToArray();
        var at = IndexOf(bytes, marker);
        Assert.True(at > 0, "the fixture's tail entry was not found in the archive");
        for (var i = at + marker.Length; i < Math.Min(at + marker.Length + 24, bytes.Length); i++)
            bytes[i] ^= 0xFF;
        File.WriteAllBytes(path, bytes);

        return path;
    }

    /// <summary>
    ///     Finds a byte sequence.
    /// </summary>
    /// <param name="haystack">Bytes to search.</param>
    /// <param name="needle">Bytes to find.</param>
    /// <returns>The index, or -1.</returns>
    private static int IndexOf(byte[] haystack, byte[] needle)
    {
        for (var i = 0; i <= haystack.Length - needle.Length; i++)
        {
            var found = true;
            for (var j = 0; j < needle.Length && found; j++)
                if (haystack[i + j] != needle[j])
                    found = false;

            if (found) return i;
        }

        return -1;
    }

    /// <summary>
    ///     A container this guard cannot read in full must be refused, not scanned in part.
    /// </summary>
    [Fact]
    public void ConvertingAnEpubWithAnUnreadableEntry_ShouldBeRefused()
    {
        var input = WriteEpub(CreateTestFilePath("partial.epub"), true);
        var output = CreateTestFilePath("partial.pdf");

        var exception = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []));

        Assert.Contains("could not be read in full", exception.Message);
        Assert.False(File.Exists(output));
    }

    /// <summary>
    ///     The refusal must not extend to a container that reads cleanly, or the guard would just
    ///     disable EPUB conversion.
    /// </summary>
    [Fact]
    public void ConvertingAReadableEpub_ShouldStillWork()
    {
        var input = WriteEpub(CreateTestFilePath("readable.epub"), false);
        var output = CreateTestFilePath("readable.pdf");

        DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []);

        Assert.True(File.Exists(output));
    }

    /// <summary>
    ///     An MHT larger than the parser limit must be refused from its size on disk, before the
    ///     MIME parser is handed it.
    /// </summary>
    [Fact]
    public void ConvertingAnOversizedMht_ShouldBeRefusedBeforeParsing()
    {
        var input = CreateTestFilePath("oversized.mht");

        // 64 MiB + a little. Written as a sparse-ish single write so the fixture stays quick.
        using (var file = new FileStream(input, FileMode.Create, FileAccess.Write))
        {
            var header = Encoding.ASCII.GetBytes(
                "From: <saved by test>\r\nSubject: oversized\r\nMIME-Version: 1.0\r\n" +
                "Content-Type: text/html\r\n\r\n<html><body>");
            file.Write(header);
            var filler = new byte[1024 * 1024];
            Array.Fill(filler, (byte)'a');
            for (var written = 0; written < 65; written++)
                file.Write(filler);
        }

        var output = CreateTestFilePath("oversized.pdf");

        var exception = Assert.Throws<ArgumentException>(() =>
            DocumentConverter.ConvertToPdfFromSpecialFormat(input, output, []));

        // Refused before it is even staged now (R23-RES01): the copy admission speaks first,
        // on the size it saw on the opened handle, and no bytes of it reach the staging area.
        Assert.Contains("above the", exception.Message);
        Assert.Contains("68,157,541", exception.Message);
        Assert.False(File.Exists(output));
    }
}
