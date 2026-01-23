using AsposeMcpServer.Handlers.Excel.Properties;
using AsposeMcpServer.Results.Excel.Properties;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWorkbookPropertiesResult>(res);
        Assert.Equal("Test Title", result.Title);
        Assert.Equal("Test Author", result.Author);
        Assert.True(result.TotalSheets >= 1);
    }

    [Fact]
    public void Execute_ReturnsAllBuiltInProperties()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.BuiltInDocumentProperties.Title = "TestTitle";
        workbook.BuiltInDocumentProperties.Subject = "TestSubject";
        workbook.BuiltInDocumentProperties.Author = "TestAuthor";
        workbook.BuiltInDocumentProperties.Keywords = "TestKeywords";
        workbook.BuiltInDocumentProperties.Comments = "TestComments";
        workbook.BuiltInDocumentProperties.Category = "TestCategory";
        workbook.BuiltInDocumentProperties.Company = "TestCompany";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWorkbookPropertiesResult>(res);
        Assert.Equal("TestTitle", result.Title);
        Assert.Equal("TestSubject", result.Subject);
        Assert.Equal("TestAuthor", result.Author);
        Assert.Equal("TestKeywords", result.Keywords);
        Assert.Equal("TestComments", result.Comments);
        Assert.Equal("TestCategory", result.Category);
        Assert.Equal("TestCompany", result.Company);
    }

    [Fact]
    public void Execute_ReturnsCustomProperties()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.CustomDocumentProperties.Add("CustomProp", "CustomValue");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWorkbookPropertiesResult>(res);
        Assert.NotNull(result.CustomProperties);
        Assert.Contains(result.CustomProperties, p => p is { Name: "CustomProp", Value: "CustomValue" });
    }

    [Fact]
    public void Execute_ReturnsTimestamps()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWorkbookPropertiesResult>(res);
        Assert.NotNull(result.Created);
        Assert.NotNull(result.Modified);
    }

    #endregion
}
