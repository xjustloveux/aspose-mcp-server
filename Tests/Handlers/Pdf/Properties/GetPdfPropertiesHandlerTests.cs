using AsposeMcpServer.Handlers.Pdf.Properties;
using AsposeMcpServer.Results.Pdf.Properties;
using AsposeMcpServer.Tests.Infrastructure;

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
    public void Execute_ReturnsTypedResult()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GetPropertiesPdfResult>(result);
    }

    [Fact]
    public void Execute_ReturnsTotalPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal(3, result.TotalPages);
    }

    [Fact]
    public void Execute_ReturnsIsEncrypted()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.False(result.IsEncrypted);
    }

    [Fact]
    public void Execute_ReturnsIsLinearized()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.False(result.IsLinearized);
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
    public void Execute_ReturnsCorrectType()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GetPropertiesPdfResult>(result);
    }

    [Fact]
    public void Execute_ReturnsTitleProperty()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Info.Title = "Test Title";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal("Test Title", result.Title);
    }

    [Fact]
    public void Execute_ReturnsAuthorProperty()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Info.Author = "Test Author";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal("Test Author", result.Author);
    }

    [Fact]
    public void Execute_ReturnsSubjectProperty()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Info.Subject = "Test Subject";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal("Test Subject", result.Subject);
    }

    [Fact]
    public void Execute_ReturnsKeywordsProperty()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Info.Keywords = "test, keywords";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal("test, keywords", result.Keywords);
    }

    [Fact]
    public void Execute_ReturnsCreatorProperty()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Info.Creator = "Test Creator";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal("Test Creator", result.Creator);
    }

    [Fact]
    public void Execute_ReturnsProducerProperty()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Info.Producer = "Test Producer";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal("Test Producer", result.Producer);
    }

    [Fact]
    public void Execute_ReturnsCreationDateProperty()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Info.CreationDate = DateTime.Now;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.NotNull(result.CreationDate);
    }

    [Fact]
    public void Execute_ReturnsModificationDateProperty()
    {
        var doc = CreateDocumentWithPages(1);
        doc.Info.ModDate = DateTime.Now;
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.NotNull(result.ModificationDate);
    }

    #endregion

    #region Various Page Counts

    [Fact]
    public void Execute_WithSinglePage_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal(1, result.TotalPages);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPdfResult>(res);
        Assert.Equal(pageCount, result.TotalPages);
    }

    #endregion
}
