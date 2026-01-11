using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Formula;

public class CalculateFormulasHandlerTests : ExcelHandlerTestBase
{
    private readonly CalculateFormulasHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Calculate()
    {
        Assert.Equal("calculate", _handler.Operation);
    }

    #endregion

    #region Empty Workbook

    [Fact]
    public void Execute_EmptyWorkbook_ReturnsSuccess()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Formulas calculated", result);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithUncalculatedFormulas()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["B1"].Value = 20;
        sheet.Cells["C1"].Formula = "=A1+B1";
        return workbook;
    }

    #endregion

    #region Basic Calculate Operations

    [Fact]
    public void Execute_CalculatesFormulas()
    {
        var workbook = CreateWorkbookWithUncalculatedFormulas();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Formulas calculated", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_UpdatesCellValues()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["B1"].Value = 20;
        sheet.Cells["C1"].Formula = "=A1+B1";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(30.0, sheet.Cells["C1"].Value);
    }

    [Fact]
    public void Execute_CalculatesMultipleFormulas()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 10;
        sheet.Cells["A2"].Value = 20;
        sheet.Cells["A3"].Value = 30;
        sheet.Cells["B1"].Formula = "=A1*2";
        sheet.Cells["B2"].Formula = "=A2*2";
        sheet.Cells["B3"].Formula = "=A3*2";
        sheet.Cells["C1"].Formula = "=SUM(B1:B3)";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(20.0, sheet.Cells["B1"].Value);
        Assert.Equal(40.0, sheet.Cells["B2"].Value);
        Assert.Equal(60.0, sheet.Cells["B3"].Value);
        Assert.Equal(120.0, sheet.Cells["C1"].Value);
    }

    [Fact]
    public void Execute_CalculatesAcrossSheets()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = 100;
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Formula = "=Sheet1!A1*2";
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(200.0, workbook.Worksheets[1].Cells["A1"].Value);
    }

    #endregion
}
