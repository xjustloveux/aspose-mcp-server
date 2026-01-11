using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class ShowFormulasExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly ShowFormulasExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ShowFormulas()
    {
        Assert.Equal("show_formulas", _handler.Operation);
    }

    #endregion

    #region Basic Show Operations

    [Fact]
    public void Execute_ShowsFormulas()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "visible", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shown", result.ToLower());
        Assert.True(workbook.Worksheets[0].ShowFormulas);
        AssertModified(context);
    }

    [Fact]
    public void Execute_HidesFormulas()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].ShowFormulas = true;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "visible", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("hidden", result.ToLower());
        Assert.False(workbook.Worksheets[0].ShowFormulas);
    }

    [Fact]
    public void Execute_WithSheetIndex_ShowsFormulasOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "visible", true }
        });

        _handler.Execute(context, parameters);

        Assert.True(workbook.Worksheets[1].ShowFormulas);
    }

    #endregion
}
