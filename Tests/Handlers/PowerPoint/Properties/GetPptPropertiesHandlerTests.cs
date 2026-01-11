using AsposeMcpServer.Handlers.PowerPoint.Properties;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Properties;

public class GetPptPropertiesHandlerTests : PptHandlerTestBase
{
    private readonly GetPptPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsDocumentProperties()
    {
        var pres = CreateEmptyPresentation();
        pres.DocumentProperties.Title = "Test Title";
        pres.DocumentProperties.Author = "Test Author";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Test Title", result);
        Assert.Contains("Test Author", result);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_GetsAllProperties()
    {
        var pres = CreateEmptyPresentation();
        pres.DocumentProperties.Title = "My Title";
        pres.DocumentProperties.Subject = "My Subject";
        pres.DocumentProperties.Author = "My Author";
        pres.DocumentProperties.Keywords = "key1, key2";
        pres.DocumentProperties.Comments = "My Comments";
        pres.DocumentProperties.Category = "My Category";
        pres.DocumentProperties.Company = "My Company";
        pres.DocumentProperties.Manager = "My Manager";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("My Title", result);
        Assert.Contains("My Subject", result);
        Assert.Contains("My Author", result);
        Assert.Contains("key1, key2", result);
        Assert.Contains("My Comments", result);
        Assert.Contains("My Category", result);
        Assert.Contains("My Company", result);
        Assert.Contains("My Manager", result);
    }

    [Fact]
    public void Execute_WithEmptyProperties_ReturnsEmptyValues()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("title", result.ToLower());
        Assert.Contains("author", result.ToLower());
    }

    [Fact]
    public void Execute_IncludesCreatedTime()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("createdTime", result);
    }

    [Fact]
    public void Execute_IncludesRevisionNumber()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("revisionNumber", result);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var pres = CreateEmptyPresentation();
        pres.DocumentProperties.Title = "JSON Test";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
        Assert.Contains("\"title\"", result);
    }

    #endregion
}
