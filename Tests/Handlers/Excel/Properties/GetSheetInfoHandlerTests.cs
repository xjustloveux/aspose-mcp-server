using AsposeMcpServer.Handlers.Excel.Properties;
using AsposeMcpServer.Results.Excel.Properties;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Properties;

public class GetSheetInfoHandlerTests : ExcelHandlerTestBase
{
    private readonly GetSheetInfoHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetSheetInfo()
    {
        Assert.Equal("get_sheet_info", _handler.Operation);
    }

    #endregion

    #region Basic Get Sheet Info Operations

    [Fact]
    public void Execute_ReturnsAllSheetsInfo()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetInfoResult>(res);
        Assert.Equal(3, result.Count);
        Assert.Equal(3, result.TotalWorksheets);
        Assert.NotNull(result.Items);
        Assert.Equal(3, result.Items.Count);
    }

    [Fact]
    public void Execute_WithTargetSheetIndex_ReturnsSpecificSheetInfo()
    {
        var workbook = CreateWorkbookWithSheets(3);
        workbook.Worksheets[1].Name = "TargetSheet";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetSheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetInfoResult>(res);
        Assert.Equal(1, result.Count);
        Assert.Single(result.Items);
        Assert.Equal("TargetSheet", result.Items[0].Name);
    }

    [Fact]
    public void Execute_ReturnsSheetDataCounts()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Data");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetInfoResult>(res);
        var sheetDetail = Assert.Single(result.Items);
        Assert.True(sheetDetail.DataRowCount >= 0);
        Assert.True(sheetDetail.DataColumnCount >= 0);
        Assert.NotNull(sheetDetail.UsedRange);
    }

    [Fact]
    public void Execute_ReturnsPageSetupInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetInfoResult>(res);
        var sheetDetail = Assert.Single(result.Items);
        Assert.NotNull(sheetDetail.PageOrientation);
        Assert.NotNull(sheetDetail.PaperSize);
    }

    [Fact]
    public void Execute_ReturnsFreezePanesInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSheetInfoResult>(res);
        var sheetDetail = Assert.Single(result.Items);
        Assert.NotNull(sheetDetail.FreezePanes);
    }

    #endregion
}
