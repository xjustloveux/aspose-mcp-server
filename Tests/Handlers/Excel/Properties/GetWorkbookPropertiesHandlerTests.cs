using AsposeMcpServer.Handlers.Excel.Properties;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Properties;

public class GetWorkbookPropertiesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetWorkbookPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetWorkbookProperties()
    {
        Assert.Equal("get_workbook_properties", _handler.Operation);
    }

    #endregion

    #region Basic Get Workbook Properties Operations

    [Fact]
    public void Execute_ReturnsWorkbookProperties()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.BuiltInDocumentProperties.Title = "Test Title";
        workbook.BuiltInDocumentProperties.Author = "Test Author";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Test Title", result);
        Assert.Contains("Test Author", result);
        Assert.Contains("totalSheets", result);
    }

    [Fact]
    public void Execute_ReturnsAllBuiltInProperties()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("title", result);
        Assert.Contains("subject", result);
        Assert.Contains("author", result);
        Assert.Contains("keywords", result);
        Assert.Contains("comments", result);
        Assert.Contains("category", result);
        Assert.Contains("company", result);
    }

    [Fact]
    public void Execute_ReturnsCustomProperties()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.CustomDocumentProperties.Add("CustomProp", "CustomValue");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("customProperties", result);
        Assert.Contains("CustomProp", result);
        Assert.Contains("CustomValue", result);
    }

    [Fact]
    public void Execute_ReturnsTimestamps()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("created", result);
        Assert.Contains("modified", result);
    }

    #endregion
}
