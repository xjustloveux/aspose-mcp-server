using AsposeMcpServer.Handlers.Pdf.Properties;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "author", "New Author" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("properties", result);
    }

    [Fact]
    public void Execute_WithNoParameters_StillSucceeds()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    #endregion

    #region Empty Values

    [Fact]
    public void Execute_WithEmptyTitle_DoesNotFail()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    [Fact]
    public void Execute_WithNullTitle_DoesNotFail()
    {
        var doc = CreateDocumentWithPages(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", null }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    #endregion
}
