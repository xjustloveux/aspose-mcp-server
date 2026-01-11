using Aspose.Words;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Base class for Word tool tests providing Word-specific functionality
/// </summary>
public abstract class WordTestBase : TestBase
{
    /// <summary>
    ///     Creates a new Word document for testing
    /// </summary>
    protected string CreateWordDocument(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var doc = new Document();
        doc.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a Word document with sample content
    /// </summary>
    protected string CreateWordDocumentWithContent(string fileName, string content)
    {
        var filePath = CreateTestFilePath(fileName);
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write(content);
        doc.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a Word document with multiple paragraphs
    /// </summary>
    protected string CreateWordDocumentWithParagraphs(string fileName, params string[] paragraphs)
    {
        var filePath = CreateTestFilePath(fileName);
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        foreach (var para in paragraphs) builder.Writeln(para);

        doc.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Verifies that a paragraph exists and has the expected text
    /// </summary>
    protected void AssertParagraphExists(Document doc, int paragraphIndex, string expectedText)
    {
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        Assert.True(paragraphIndex < paragraphs.Count, $"Paragraph index {paragraphIndex} is out of range");
        Assert.Equal(expectedText, paragraphs[paragraphIndex].GetText().Trim());
    }

    /// <summary>
    ///     Verifies that a paragraph has the expected style
    /// </summary>
    protected void AssertParagraphStyle(Document doc, int paragraphIndex, string expectedStyleName)
    {
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        Assert.True(paragraphIndex < paragraphs.Count, $"Paragraph index {paragraphIndex} is out of range");
        Assert.Equal(expectedStyleName, paragraphs[paragraphIndex].ParagraphFormat.StyleName);
    }

    /// <summary>
    ///     Gets all paragraphs from a document
    /// </summary>
    protected List<Paragraph> GetParagraphs(Document doc, bool includeNested = true)
    {
        return doc.GetChildNodes(NodeType.Paragraph, includeNested).Cast<Paragraph>().ToList();
    }

    /// <summary>
    ///     Checks if Aspose.Words is running in evaluation mode.
    /// </summary>
    protected new static bool IsEvaluationMode(AsposeLibraryType libraryType = AsposeLibraryType.Words)
    {
        return TestBase.IsEvaluationMode(libraryType);
    }
}
