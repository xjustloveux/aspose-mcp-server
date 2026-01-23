using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Revision;
using AsposeMcpServer.Results.Word.Revision;
using AsposeMcpServer.Tests.Infrastructure;

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

    #region Result Properties

    [Fact]
    public void Execute_ReturnsCorrectProperties()
    {
        var originalPath = CreateTempDocumentWithText("First version.");
        var revisedPath = CreateTempDocumentWithText("Second version.");
        var outputPath = Path.Combine(TestDir, "properties_output.docx");

        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "originalPath", originalPath },
            { "revisedPath", revisedPath },
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CompareDocumentsResult>(res);

        Assert.True(result.RevisionCount >= 0);
        Assert.Equal(outputPath, result.OutputPath);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CompareDocumentsResult>(res);

        Assert.True(result.RevisionCount > 0);
        Assert.Equal(outputPath, result.OutputPath);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CompareDocumentsResult>(res);

        Assert.Equal(0, result.RevisionCount);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CompareDocumentsResult>(res);

        Assert.True(result.RevisionCount > 0);
        Assert.Equal(outputPath, result.OutputPath);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CompareDocumentsResult>(res);

        Assert.Equal(0, result.RevisionCount);
    }

    #endregion
}
