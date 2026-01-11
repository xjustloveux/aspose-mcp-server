using AsposeMcpServer.Handlers.Pdf.Properties;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Properties;

public class GetPdfPropertiesHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Basic Properties Retrieval

    [Fact]
    public void Execute_ReturnsJsonResult()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
    }

    [Fact]
    public void Execute_ReturnsTotalPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"totalPages\": 3", result);
    }

    [Fact]
    public void Execute_ReturnsIsEncrypted()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"isEncrypted\":", result);
    }

    [Fact]
    public void Execute_ReturnsIsLinearized()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"isLinearized\":", result);
    }

    [Fact]
    public void Execute_DoesNotMarkAsModified()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.False(context.IsModified);
    }

    #endregion

    #region Metadata Properties

    [Fact]
    public void Execute_ReturnsTitleProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"title\":", result);
    }

    [Fact]
    public void Execute_ReturnsAuthorProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"author\":", result);
    }

    [Fact]
    public void Execute_ReturnsSubjectProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"subject\":", result);
    }

    [Fact]
    public void Execute_ReturnsKeywordsProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"keywords\":", result);
    }

    [Fact]
    public void Execute_ReturnsCreatorProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"creator\":", result);
    }

    [Fact]
    public void Execute_ReturnsProducerProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"producer\":", result);
    }

    [Fact]
    public void Execute_ReturnsCreationDateProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"creationDate\":", result);
    }

    [Fact]
    public void Execute_ReturnsModificationDateProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"modificationDate\":", result);
    }

    #endregion

    #region Various Page Counts

    [Fact]
    public void Execute_WithSinglePage_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"totalPages\": 1", result);
    }

    [SkippableTheory]
    [InlineData(5)]
    [InlineData(10)]
    public void Execute_WithHighPageCounts_ReturnsCorrectCount(int pageCount)
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf, "PDF evaluation mode limits page count to 4");

        var doc = CreateDocumentWithPages(pageCount);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"\"totalPages\": {pageCount}", result);
    }

    #endregion
}
