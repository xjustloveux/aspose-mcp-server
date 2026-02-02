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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("New Title", pres.DocumentProperties.Title);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("New Subject", pres.DocumentProperties.Subject);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("New Author", pres.DocumentProperties.Author);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("keyword1, keyword2", pres.DocumentProperties.Keywords);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("New Comments", pres.DocumentProperties.Comments);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("New Category", pres.DocumentProperties.Category);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("New Company", pres.DocumentProperties.Company);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("New Manager", pres.DocumentProperties.Manager);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            Assert.Equal("Multi Title", pres.DocumentProperties.Title);
            Assert.Equal("Multi Author", pres.DocumentProperties.Author);
            Assert.Equal("Multi Company", pres.DocumentProperties.Company);
        }

        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal("CustomValue", pres.DocumentProperties["CustomKey"]);
        AssertModified(context);
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

        Assert.IsType<SuccessResult>(res);
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
