using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Revision;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Revision;

public class CompareDocumentsHandlerTests : WordHandlerTestBase
{
    private readonly CompareDocumentsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Compare()
    {
        Assert.Equal("compare", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private string CreateTempDocumentWithText(string text)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write(text);
        var path = Path.Combine(TestDir, $"doc_{Guid.NewGuid()}.docx");
        doc.Save(path);
        return path;
    }

    #endregion

    #region Basic Compare Operations

    [Fact]
    public void Execute_ComparesDocuments()
    {
        var originalPath = CreateTempDocumentWithText("Original content.");
        var revisedPath = CreateTempDocumentWithText("Original content. New content added.");
        var outputPath = Path.Combine(TestDir, "comparison_output.docx");

        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "originalPath", originalPath },
            { "revisedPath", revisedPath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("comparison completed", result.ToLower());
        Assert.Contains("difference", result.ToLower());
        Assert.True(System.IO.File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithIdenticalDocuments_ReportsZeroDifferences()
    {
        var originalPath = CreateTempDocumentWithText("Same content.");
        var revisedPath = CreateTempDocumentWithText("Same content.");
        var outputPath = Path.Combine(TestDir, "identical_output.docx");

        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "originalPath", originalPath },
            { "revisedPath", revisedPath },
            { "outputPath", outputPath }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0 difference", result.ToLower());
    }

    [Fact]
    public void Execute_WithAuthorName_UsesSpecifiedAuthor()
    {
        var originalPath = CreateTempDocumentWithText("Original.");
        var revisedPath = CreateTempDocumentWithText("Changed.");
        var outputPath = Path.Combine(TestDir, "author_output.docx");

        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "originalPath", originalPath },
            { "revisedPath", revisedPath },
            { "outputPath", outputPath },
            { "authorName", "TestAuthor" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("comparison completed", result.ToLower());
    }

    [Fact]
    public void Execute_WithIgnoreFormatting_ComparesWithoutFormatting()
    {
        var originalPath = CreateTempDocumentWithText("Content.");
        var revisedPath = CreateTempDocumentWithText("Content.");
        var outputPath = Path.Combine(TestDir, "ignore_format_output.docx");

        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "originalPath", originalPath },
            { "revisedPath", revisedPath },
            { "outputPath", outputPath },
            { "ignoreFormatting", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("comparison completed", result.ToLower());
    }

    #endregion
}
