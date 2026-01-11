using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Content;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Content;

public class GetWordDocumentInfoHandlerTests : WordHandlerTestBase
{
    private readonly GetWordDocumentInfoHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetDocumentInfo()
    {
        Assert.Equal("get_document_info", _handler.Operation);
    }

    #endregion

    #region Multiple Sections

    [Fact]
    public void Execute_WithMultipleSections_ReturnsCorrectCount()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"sections\": 3", result);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTabStops()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.ParagraphFormat.TabStops.Add(72, TabAlignment.Left, TabLeader.None);
        builder.Writeln("Text with tab stop");
        return doc;
    }

    #endregion

    #region Basic Info Retrieval

    [Fact]
    public void Execute_ReturnsJsonResult()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
    }

    [Fact]
    public void Execute_ReturnsSectionsCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"sections\":", result);
    }

    [Fact]
    public void Execute_ReturnsCreatedDate()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"created\":", result);
    }

    [Fact]
    public void Execute_ReturnsModifiedDate()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"modified\":", result);
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

    #endregion

    #region Document Properties

    [Fact]
    public void Execute_ReturnsTitleProperty()
    {
        var doc = CreateDocumentWithText("Test");
        doc.BuiltInDocumentProperties.Title = "Test Title";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"title\":", result);
    }

    [Fact]
    public void Execute_ReturnsAuthorProperty()
    {
        var doc = CreateDocumentWithText("Test");
        doc.BuiltInDocumentProperties.Author = "Test Author";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"author\":", result);
    }

    [Fact]
    public void Execute_ReturnsSubjectProperty()
    {
        var doc = CreateDocumentWithText("Test");
        doc.BuiltInDocumentProperties.Subject = "Test Subject";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"subject\":", result);
    }

    [Fact]
    public void Execute_ReturnsPagesCount()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"pages\":", result);
    }

    #endregion

    #region Include TabStops Parameter

    [Fact]
    public void Execute_WithIncludeTabStopsTrue_ReturnsTabStopsInfo()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeTabStops", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"tabStopsIncluded\": true", result);
    }

    [Fact]
    public void Execute_WithIncludeTabStopsFalse_ShowsFalse()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeTabStops", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"tabStopsIncluded\": false", result);
    }

    [Fact]
    public void Execute_DefaultIncludeTabStops_IsFalse()
    {
        var doc = CreateDocumentWithText("Test");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"tabStopsIncluded\": false", result);
    }

    [Fact]
    public void Execute_WithTabStops_ReturnsTabStopsArray()
    {
        var doc = CreateDocumentWithTabStops();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "includeTabStops", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"tabStops\":", result);
    }

    #endregion
}
