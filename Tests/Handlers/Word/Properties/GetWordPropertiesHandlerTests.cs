using AsposeMcpServer.Handlers.Word.Properties;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Properties;

public class GetWordPropertiesHandlerTests : WordHandlerTestBase
{
    private readonly GetWordPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Basic Get Properties Operations

    [Fact]
    public void Execute_ReturnsBuiltInProperties()
    {
        var doc = CreateEmptyDocument();
        doc.BuiltInDocumentProperties.Title = "Test Title";
        doc.BuiltInDocumentProperties.Author = "Test Author";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("builtInProperties", result);
        Assert.Contains("Test Title", result);
        Assert.Contains("Test Author", result);
    }

    [Fact]
    public void Execute_ReturnsStatistics()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("statistics", result);
        Assert.Contains("wordCount", result);
        Assert.Contains("characterCount", result);
        Assert.Contains("pageCount", result);
    }

    [Fact]
    public void Execute_ReturnsCustomProperties_WhenPresent()
    {
        var doc = CreateEmptyDocument();
        doc.CustomDocumentProperties.Add("CustomProp", "CustomValue");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("customProperties", result);
        Assert.Contains("CustomProp", result);
        Assert.Contains("CustomValue", result);
    }

    [Fact]
    public void Execute_WithNoCustomProperties_DoesNotIncludeCustomProperties()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("builtInProperties", result);
        Assert.DoesNotContain("customProperties", result);
    }

    [Fact]
    public void Execute_ReturnsAllBuiltInFields()
    {
        var doc = CreateEmptyDocument();
        doc.BuiltInDocumentProperties.Subject = "Test Subject";
        doc.BuiltInDocumentProperties.Keywords = "test, keywords";
        doc.BuiltInDocumentProperties.Category = "Test Category";
        doc.BuiltInDocumentProperties.Company = "Test Company";
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("subject", result);
        Assert.Contains("keywords", result);
        Assert.Contains("category", result);
        Assert.Contains("company", result);
        Assert.Contains("createdTime", result);
    }

    #endregion
}
