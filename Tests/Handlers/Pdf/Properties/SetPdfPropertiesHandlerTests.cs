using AsposeMcpServer.Handlers.Pdf.Properties;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Properties;

public class SetPdfPropertiesHandlerTests : PdfHandlerTestBase
{
    private readonly SetPdfPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Set()
    {
        Assert.Equal("set", _handler.Operation);
    }

    #endregion

    #region Set Author

    [Fact]
    public void Execute_WithAuthor_SetsAuthor()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "author", "John Doe" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal("John Doe", doc.Info.Author);
        AssertModified(context);
    }

    #endregion

    #region Set Subject

    [Fact]
    public void Execute_WithSubject_SetsSubject()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "subject", "Test Subject" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal("Test Subject", doc.Info.Subject);
        AssertModified(context);
    }

    #endregion

    #region Set Keywords

    [Fact]
    public void Execute_WithKeywords_SetsKeywords()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "keywords", "test, pdf, keywords" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal("test, pdf, keywords", doc.Info.Keywords);
        AssertModified(context);
    }

    #endregion

    #region Set Creator

    [Fact]
    public void Execute_WithCreator_SetsCreator()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "creator", "Test Application" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal("Test Application", doc.Info.Creator);
        AssertModified(context);
    }

    #endregion

    #region Set Producer

    [Fact]
    public void Execute_WithProducer_SetsProducer()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "producer", "Test Producer" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal("Test Producer", doc.Info.Producer);
        AssertModified(context);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsProperties()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "New Title" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal("New Title", doc.Info.Title);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsAuthorProperty()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "author", "New Author" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal("New Author", doc.Info.Author);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoParameters_StillSucceeds()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion

    #region Set Title

    [Fact]
    public void Execute_WithTitle_SetsTitle()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Document Title" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal("Document Title", doc.Info.Title);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithTitle_MarksModified()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "New Title" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Multiple Properties

    [Fact]
    public void Execute_WithMultipleProperties_SetsAll()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "My Document" },
            { "author", "Jane Smith" },
            { "subject", "Document Subject" },
            { "keywords", "key1, key2, key3" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            Assert.Equal("My Document", doc.Info.Title);
            Assert.Equal("Jane Smith", doc.Info.Author);
            Assert.Equal("Document Subject", doc.Info.Subject);
            Assert.Equal("key1, key2, key3", doc.Info.Keywords);
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithAllProperties_SetsAll()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Full Title" },
            { "author", "Full Author" },
            { "subject", "Full Subject" },
            { "keywords", "Full Keywords" },
            { "creator", "Full Creator" },
            { "producer", "Full Producer" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            Assert.Equal("Full Title", doc.Info.Title);
            Assert.Equal("Full Author", doc.Info.Author);
            Assert.Equal("Full Subject", doc.Info.Subject);
            Assert.Equal("Full Keywords", doc.Info.Keywords);
            Assert.Equal("Full Creator", doc.Info.Creator);
            Assert.Equal("Full Producer", doc.Info.Producer);
        }

        AssertModified(context);
    }

    #endregion

    #region Empty Values

    [Fact]
    public void Execute_WithEmptyTitle_DoesNotFail()
    {
        var doc = CreateDocumentWithPages(1);
        var originalTitle = doc.Info.Title;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal(originalTitle, doc.Info.Title);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNullTitle_DoesNotFail()
    {
        var doc = CreateDocumentWithPages(1);
        var originalTitle = doc.Info.Title;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", null }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
            Assert.Equal(originalTitle, doc.Info.Title);
        AssertModified(context);
    }

    #endregion
}
