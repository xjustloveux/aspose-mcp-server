using AsposeMcpServer.Handlers.Excel.Properties;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Properties;

public class GetSheetPropertiesHandlerTests : ExcelHandlerTestBase
{
    private readonly GetSheetPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetSheetProperties()
    {
        Assert.Equal("get_sheet_properties", _handler.Operation);
    }

    #endregion

    #region Basic Get Sheet Properties Operations

    [Fact]
    public void Execute_ReturnsSheetProperties()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Name = "TestSheet";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("TestSheet", result);
        Assert.Contains("\"index\"", result);
        Assert.Contains("isVisible", result);
    }

    [Fact]
    public void Execute_WithSheetIndex_ReturnsSpecificSheetProperties()
    {
        var workbook = CreateWorkbookWithSheets(2);
        workbook.Worksheets[1].Name = "SecondSheet";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("SecondSheet", result);
    }

    [Fact]
    public void Execute_ReturnsDataCounts()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Data");
        workbook.Worksheets[0].Cells["B2"].PutValue("More Data");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("dataRowCount", result);
        Assert.Contains("dataColumnCount", result);
    }

    [Fact]
    public void Execute_ReturnsPrintSettings()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("printSettings", result);
        Assert.Contains("printArea", result);
        Assert.Contains("orientation", result);
        Assert.Contains("paperSize", result);
    }

    [Fact]
    public void Execute_ReturnsObjectCounts()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("commentsCount", result);
        Assert.Contains("chartsCount", result);
        Assert.Contains("picturesCount", result);
        Assert.Contains("hyperlinksCount", result);
    }

    #endregion
}
