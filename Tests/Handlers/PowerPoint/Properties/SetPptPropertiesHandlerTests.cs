using AsposeMcpServer.Handlers.PowerPoint.Properties;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Properties;

public class SetPptPropertiesHandlerTests : PptHandlerTestBase
{
    private readonly SetPptPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Set()
    {
        Assert.Equal("set", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsTitle()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "New Title" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Title", result.Message);
        Assert.Equal("New Title", pres.DocumentProperties.Title);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsSubject()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "subject", "New Subject" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Subject", result.Message);
        Assert.Equal("New Subject", pres.DocumentProperties.Subject);
    }

    [Fact]
    public void Execute_SetsAuthor()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "author", "New Author" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Author", result.Message);
        Assert.Equal("New Author", pres.DocumentProperties.Author);
    }

    [Fact]
    public void Execute_SetsKeywords()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "keywords", "keyword1, keyword2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Keywords", result.Message);
        Assert.Equal("keyword1, keyword2", pres.DocumentProperties.Keywords);
    }

    [Fact]
    public void Execute_SetsComments()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "comments", "New Comments" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Comments", result.Message);
        Assert.Equal("New Comments", pres.DocumentProperties.Comments);
    }

    [Fact]
    public void Execute_SetsCategory()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "category", "New Category" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Category", result.Message);
        Assert.Equal("New Category", pres.DocumentProperties.Category);
    }

    [Fact]
    public void Execute_SetsCompany()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "company", "New Company" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Company", result.Message);
        Assert.Equal("New Company", pres.DocumentProperties.Company);
    }

    [Fact]
    public void Execute_SetsManager()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "manager", "New Manager" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Manager", result.Message);
        Assert.Equal("New Manager", pres.DocumentProperties.Manager);
    }

    [Fact]
    public void Execute_SetsMultipleProperties()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Multi Title" },
            { "author", "Multi Author" },
            { "company", "Multi Company" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Title", result.Message);
        Assert.Contains("Author", result.Message);
        Assert.Contains("Company", result.Message);
        Assert.Equal("Multi Title", pres.DocumentProperties.Title);
        Assert.Equal("Multi Author", pres.DocumentProperties.Author);
        Assert.Equal("Multi Company", pres.DocumentProperties.Company);
    }

    [Fact]
    public void Execute_SetsCustomProperties()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var customProps = new Dictionary<string, object>
        {
            { "CustomKey", "CustomValue" }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "customProperties", customProps }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("CustomProperties", result.Message);
        Assert.Equal("CustomValue", pres.DocumentProperties["CustomKey"]);
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void Execute_WithEmptyParameters_StillReturnsSuccess()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_OverwritesExistingProperties()
    {
        var pres = CreateEmptyPresentation();
        pres.DocumentProperties.Title = "Old Title";
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "New Title" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("New Title", pres.DocumentProperties.Title);
    }

    #endregion
}
