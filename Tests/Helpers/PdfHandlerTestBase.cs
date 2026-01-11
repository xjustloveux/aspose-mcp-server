using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Base class for PDF Handler tests providing PDF-specific test infrastructure.
/// </summary>
public abstract class PdfHandlerTestBase : HandlerTestBase<Document>
{
    /// <summary>
    ///     Creates a new empty PDF document for testing.
    /// </summary>
    /// <returns>A new empty Document instance with one page.</returns>
    protected static Document CreateEmptyDocument()
    {
        var doc = new Document();
        doc.Pages.Add();
        return doc;
    }

    /// <summary>
    ///     Creates a PDF document with text content.
    /// </summary>
    /// <param name="text">The text content to add.</param>
    /// <returns>A Document with the specified text.</returns>
    protected static Document CreateDocumentWithText(string text)
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        var textFragment = new TextFragment(text);
        page.Paragraphs.Add(textFragment);
        return doc;
    }

    /// <summary>
    ///     Creates a PDF document with multiple pages.
    /// </summary>
    /// <param name="pageCount">The number of pages to create.</param>
    /// <returns>A Document with the specified number of pages.</returns>
    protected static Document CreateDocumentWithPages(int pageCount)
    {
        var doc = new Document();
        for (var i = 0; i < pageCount; i++)
            doc.Pages.Add();
        return doc;
    }

    /// <summary>
    ///     Asserts that the document has the expected number of pages.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="expectedCount">The expected page count.</param>
    protected static void AssertPageCount(Document doc, int expectedCount)
    {
        Assert.Equal(expectedCount, doc.Pages.Count);
    }
}
