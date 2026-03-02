using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Properties;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Properties;

[SupportedOSPlatform("windows")]
public class SetPptPropertiesHandlerTests : PptHandlerTestBase
{
    private readonly SetPptPropertiesHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Set()
    {
        SkipIfNotWindows();
        Assert.Equal("set", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [SkippableFact]
    public void Execute_SetsTitle()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsSubject()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsAuthor()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsKeywords()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsComments()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsCategory()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsCompany()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsManager()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsMultipleProperties()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_SetsCustomProperties()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithEmptyParameters_StillReturnsSuccess()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_OverwritesExistingProperties()
    {
        SkipIfNotWindows();
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
