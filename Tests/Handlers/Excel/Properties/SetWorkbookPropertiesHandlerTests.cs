using AsposeMcpServer.Handlers.Excel.Properties;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Properties;

public class SetWorkbookPropertiesHandlerTests : ExcelHandlerTestBase
{
    private readonly SetWorkbookPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetWorkbookProperties()
    {
        Assert.Equal("set_workbook_properties", _handler.Operation);
    }

    #endregion

    #region Basic Set Workbook Properties Operations

    [Fact]
    public void Execute_SetsTitle()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "New Title" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated successfully", result.ToLower());
        Assert.Equal("New Title", workbook.BuiltInDocumentProperties.Title);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsMultipleProperties()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "title", "Test Title" },
            { "author", "Test Author" },
            { "subject", "Test Subject" },
            { "keywords", "test, keywords" },
            { "company", "Test Company" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Test Title", workbook.BuiltInDocumentProperties.Title);
        Assert.Equal("Test Author", workbook.BuiltInDocumentProperties.Author);
        Assert.Equal("Test Subject", workbook.BuiltInDocumentProperties.Subject);
        Assert.Equal("test, keywords", workbook.BuiltInDocumentProperties.Keywords);
        Assert.Equal("Test Company", workbook.BuiltInDocumentProperties.Company);
    }

    [Fact]
    public void Execute_SetsCustomProperties()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "customProperties", "{\"MyProp\": \"MyValue\"}" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated successfully", result.ToLower());
        var customProp = workbook.CustomDocumentProperties["MyProp"];
        Assert.NotNull(customProp);
    }

    [Fact]
    public void Execute_WithInvalidCustomPropertiesJson_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "customProperties", "invalid json" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
