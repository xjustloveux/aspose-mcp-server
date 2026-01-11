using AsposeMcpServer.Handlers.PowerPoint.Properties;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Title", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Subject", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Author", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Keywords", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Comments", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Category", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Company", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Manager", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Title", result);
        Assert.Contains("Author", result);
        Assert.Contains("Company", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("CustomProperties", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result.ToLower());
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
