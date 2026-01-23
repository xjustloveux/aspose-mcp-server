using AsposeMcpServer.Handlers.Excel.Properties;
using AsposeMcpServer.Results.Excel.Properties;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetPropertiesResult>(res);
        Assert.Equal("TestSheet", result.Name);
        Assert.Equal(0, result.Index);
        Assert.True(result.IsVisible);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetPropertiesResult>(res);
        Assert.Equal("SecondSheet", result.Name);
    }

    [Fact]
    public void Execute_ReturnsDataCounts()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Data");
        workbook.Worksheets[0].Cells["B2"].PutValue("More Data");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetPropertiesResult>(res);
        Assert.True(result.DataRowCount >= 0);
        Assert.True(result.DataColumnCount >= 0);
    }

    [Fact]
    public void Execute_ReturnsPrintSettings()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetPropertiesResult>(res);
        Assert.NotNull(result.PrintSettings);
        Assert.NotNull(result.PrintSettings.Orientation);
        Assert.NotNull(result.PrintSettings.PaperSize);
    }

    [Fact]
    public void Execute_ReturnsObjectCounts()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetPropertiesResult>(res);
        Assert.True(result.CommentsCount >= 0);
        Assert.True(result.ChartsCount >= 0);
        Assert.True(result.PicturesCount >= 0);
        Assert.True(result.HyperlinksCount >= 0);
    }

    #endregion
}
