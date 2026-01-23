using AsposeMcpServer.Handlers.Word.Properties;
using AsposeMcpServer.Results.Word.Properties;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordPropertiesResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.BuiltInProperties);
        Assert.Equal("Test Title", result.BuiltInProperties.Title);
        Assert.Equal("Test Author", result.BuiltInProperties.Author);
    }

    [Fact]
    public void Execute_ReturnsStatistics()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordPropertiesResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.Statistics);
        Assert.True(result.Statistics.WordCount >= 0);
        Assert.True(result.Statistics.CharacterCount >= 0);
        Assert.True(result.Statistics.PageCount >= 0);
    }

    [Fact]
    public void Execute_ReturnsCustomProperties_WhenPresent()
    {
        var doc = CreateEmptyDocument();
        doc.CustomDocumentProperties.Add("CustomProp", "CustomValue");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordPropertiesResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.CustomProperties);
        Assert.True(result.CustomProperties.ContainsKey("CustomProp"));
        Assert.Equal("CustomValue", result.CustomProperties["CustomProp"].Value);
    }

    [Fact]
    public void Execute_WithNoCustomProperties_DoesNotIncludeCustomProperties()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordPropertiesResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.BuiltInProperties);
        Assert.Null(result.CustomProperties);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordPropertiesResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.BuiltInProperties);
        Assert.Equal("Test Subject", result.BuiltInProperties.Subject);
        Assert.Equal("test, keywords", result.BuiltInProperties.Keywords);
        Assert.Equal("Test Category", result.BuiltInProperties.Category);
        Assert.Equal("Test Company", result.BuiltInProperties.Company);
        Assert.NotNull(result.BuiltInProperties.CreatedTime);
    }

    #endregion
}
