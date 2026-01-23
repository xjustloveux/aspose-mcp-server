using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Content;
using AsposeMcpServer.Results.Word.Content;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Content;

public class GetWordStatisticsHandlerTests : WordHandlerTestBase
{
    private readonly GetWordStatisticsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetStatistics()
    {
        Assert.Equal("get_statistics", _handler.Operation);
    }

    #endregion

    #region Basic Statistics Retrieval

    [Fact]
    public void Execute_ReturnsResult()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
    }

    [Fact]
    public void Execute_ReturnsPagesCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Pages >= 1);
    }

    [Fact]
    public void Execute_ReturnsWordsCount()
    {
        var doc = CreateDocumentWithText("One two three four five");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Words >= 5);
    }

    [Fact]
    public void Execute_ReturnsCharactersCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Characters >= 4);
    }

    [Fact]
    public void Execute_ReturnsCharactersWithSpacesCount()
    {
        var doc = CreateDocumentWithText("Test content");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.CharactersWithSpaces >= 12);
    }

    [Fact]
    public void Execute_ReturnsParagraphsCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Paragraphs >= 1);
    }

    [Fact]
    public void Execute_ReturnsLinesCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Lines >= 1);
    }

    [Fact]
    public void Execute_ReturnsTablesCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.Tables);
    }

    [Fact]
    public void Execute_ReturnsImagesCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.Images);
    }

    [Fact]
    public void Execute_ReturnsShapesCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.Shapes);
    }

    [Fact]
    public void Execute_DoesNotMarkAsModified()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.False(context.IsModified);
    }

    [Fact]
    public void Execute_ReturnsStatisticsUpdatedFlag()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.StatisticsUpdated);
    }

    #endregion

    #region Include Footnotes Parameter

    [Fact]
    public void Execute_WithIncludeFootnotesTrue_ReturnsFootnotesCount()
    {
        var doc = CreateDocumentWithFootnote();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeFootnotes", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.Footnotes);
        Assert.True(result.FootnotesIncluded);
    }

    [Fact]
    public void Execute_WithIncludeFootnotesFalse_ReturnsNullFootnotes()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeFootnotes", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.False(result.FootnotesIncluded);
    }

    [Fact]
    public void Execute_DefaultIncludeFootnotes_IsTrue()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.FootnotesIncluded);
    }

    #endregion

    #region Document With Content

    [Fact]
    public void Execute_WithTable_CountsTable()
    {
        var doc = CreateDocumentWithTable();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.Equal(1, result.Tables);
    }

    [Fact]
    public void Execute_WithMultipleParagraphs_CountsParagraphs()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordStatisticsResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Paragraphs >= 3);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFootnote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with footnote");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote text");
        return doc;
    }

    private static Document CreateDocumentWithTable()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();
        return doc;
    }

    #endregion
}
