using AsposeMcpServer.Handlers.Excel.Properties;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 3", result);
        Assert.Contains("\"totalWorksheets\": 3", result);
        Assert.Contains("items", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 1", result);
        Assert.Contains("TargetSheet", result);
    }

    [Fact]
    public void Execute_ReturnsSheetDataCounts()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Data");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("dataRowCount", result);
        Assert.Contains("dataColumnCount", result);
        Assert.Contains("usedRange", result);
    }

    [Fact]
    public void Execute_ReturnsPageSetupInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("pageOrientation", result);
        Assert.Contains("paperSize", result);
    }

    [Fact]
    public void Execute_ReturnsFreezePanesInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("freezePanes", result);
    }

    #endregion
}
