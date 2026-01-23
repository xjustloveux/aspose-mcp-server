using AsposeMcpServer.Handlers.Word.Properties;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Properties;

public class SetWordPropertiesHandlerTests : WordHandlerTestBase
{
    private readonly SetWordPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Set()
    {
        Assert.Equal("set", _handler.Operation);
    }

    #endregion

    #region Basic Set Properties Operations

    [Fact]
    public void Execute_SetsTitle()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "New Title" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("New Title", doc.BuiltInDocumentProperties.Title);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsMultipleProperties()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Test Title" },
            { "author", "Test Author" },
            { "subject", "Test Subject" },
            { "keywords", "test, keywords" },
            { "company", "Test Company" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Test Title", doc.BuiltInDocumentProperties.Title);
        Assert.Equal("Test Author", doc.BuiltInDocumentProperties.Author);
        Assert.Equal("Test Subject", doc.BuiltInDocumentProperties.Subject);
        Assert.Equal("test, keywords", doc.BuiltInDocumentProperties.Keywords);
        Assert.Equal("Test Company", doc.BuiltInDocumentProperties.Company);
    }

    [Fact]
    public void Execute_SetsComments()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "comments", "Test Comments" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Test Comments", doc.BuiltInDocumentProperties.Comments);
    }

    [Fact]
    public void Execute_SetsCategory()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "category", "Test Category" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Test Category", doc.BuiltInDocumentProperties.Category);
    }

    [Fact]
    public void Execute_SetsManager()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "manager", "Test Manager" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Test Manager", doc.BuiltInDocumentProperties.Manager);
    }

    #endregion

    #region Custom Properties

    [Fact]
    public void Execute_SetsCustomPropertiesString()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "customProperties", "{\"MyProp\": \"MyValue\"}" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(doc.CustomDocumentProperties["MyProp"]);
    }

    [Fact]
    public void Execute_SetsCustomPropertiesInt()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "customProperties", "{\"IntProp\": 42}" }
        });

        _handler.Execute(context, parameters);

        var prop = doc.CustomDocumentProperties["IntProp"];
        Assert.NotNull(prop);
    }

    [Fact]
    public void Execute_SetsCustomPropertiesBool()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "customProperties", "{\"BoolProp\": true}" }
        });

        _handler.Execute(context, parameters);

        var prop = doc.CustomDocumentProperties["BoolProp"];
        Assert.NotNull(prop);
    }

    [Fact]
    public void Execute_OverwritesExistingCustomProperty()
    {
        var doc = CreateEmptyDocument();
        doc.CustomDocumentProperties.Add("ExistingProp", "OldValue");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "customProperties", "{\"ExistingProp\": \"NewValue\"}" }
        });

        _handler.Execute(context, parameters);

        var prop = doc.CustomDocumentProperties["ExistingProp"];
        Assert.NotNull(prop);
        Assert.Equal("NewValue", prop.Value?.ToString());
    }

    #endregion
}
