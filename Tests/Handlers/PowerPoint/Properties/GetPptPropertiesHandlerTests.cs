using AsposeMcpServer.Handlers.PowerPoint.Properties;
using AsposeMcpServer.Results.PowerPoint.Properties;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPptResult>(res);

        Assert.Equal("Test Title", result.Title);
        Assert.Equal("Test Author", result.Author);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPptResult>(res);

        Assert.Equal("My Title", result.Title);
        Assert.Equal("My Subject", result.Subject);
        Assert.Equal("My Author", result.Author);
        Assert.Equal("key1, key2", result.Keywords);
        Assert.Equal("My Comments", result.Comments);
        Assert.Equal("My Category", result.Category);
        Assert.Equal("My Company", result.Company);
        Assert.Equal("My Manager", result.Manager);
    }

    [Fact]
    public void Execute_WithEmptyProperties_ReturnsEmptyValues()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPptResult>(res);

        Assert.NotNull(result);
        Assert.IsType<GetPropertiesPptResult>(result);
    }

    [Fact]
    public void Execute_IncludesCreatedTime()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPptResult>(res);

        Assert.IsType<DateTime>(result.CreatedTime);
    }

    [Fact]
    public void Execute_IncludesRevisionNumber()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPptResult>(res);

        Assert.IsType<int>(result.RevisionNumber);
    }

    [Fact]
    public void Execute_ReturnsResultType()
    {
        var pres = CreateEmptyPresentation();
        pres.DocumentProperties.Title = "JSON Test";
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPropertiesPptResult>(res);

        Assert.NotNull(result);
        Assert.IsType<GetPropertiesPptResult>(result);
        Assert.Equal("JSON Test", result.Title);
    }

    #endregion
}
