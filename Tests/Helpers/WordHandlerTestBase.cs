using Aspose.Words;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Base class for Word Handler tests providing Word-specific test infrastructure.
/// </summary>
public abstract class WordHandlerTestBase : HandlerTestBase<Document>
{
    /// <summary>
    ///     Creates a new empty Word document for testing.
    /// </summary>
    /// <returns>A new empty Document instance.</returns>
    protected static Document CreateEmptyDocument()
    {
        return new Document();
    }

    /// <summary>
    ///     Creates a Word document with initial text content.
    /// </summary>
    /// <param name="text">The initial text content.</param>
    /// <returns>A Document with the specified text.</returns>
    protected static Document CreateDocumentWithText(string text)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write(text);
        return doc;
    }

    /// <summary>
    ///     Creates a Word document with multiple paragraphs.
    /// </summary>
    /// <param name="paragraphs">The paragraph texts.</param>
    /// <returns>A Document with the specified paragraphs.</returns>
    protected static Document CreateDocumentWithParagraphs(params string[] paragraphs)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (var i = 0; i < paragraphs.Length; i++)
        {
            builder.Write(paragraphs[i]);
            if (i < paragraphs.Length - 1)
                builder.InsertParagraph();
        }

        return doc;
    }

    /// <summary>
    ///     Gets the full text content of a document.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <returns>The document text.</returns>
    protected static string GetDocumentText(Document doc)
    {
        return doc.GetText();
    }

    /// <summary>
    ///     Asserts that the document contains the specified text.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="expectedText">The expected text.</param>
    protected static void AssertContainsText(Document doc, string expectedText)
    {
        var text = doc.GetText();
        Assert.Contains(expectedText, text);
    }

    /// <summary>
    ///     Asserts that the document does not contain the specified text.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="unexpectedText">The text that should not be present.</param>
    protected static void AssertDoesNotContainText(Document doc, string unexpectedText)
    {
        var text = doc.GetText();
        Assert.DoesNotContain(unexpectedText, text);
    }
}
