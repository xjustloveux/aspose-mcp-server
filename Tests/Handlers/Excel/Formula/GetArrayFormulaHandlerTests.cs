using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Results.Excel.Formula;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Formula;

public class GetArrayFormulaHandlerTests : ExcelHandlerTestBase
{
    private readonly GetArrayFormulaHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetArray()
    {
        Assert.Equal("get_array", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_GetsFromCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = 10;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "sheetIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetArrayFormulaResult>(res);

        Assert.Equal("A1", result.Cell);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithArrayFormula()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = 1;
        sheet.Cells["A2"].Value = 2;
        sheet.Cells["A3"].Value = 3;
        sheet.Cells["B1"].Value = 10;
        sheet.Cells["B2"].Value = 20;
        sheet.Cells["B3"].Value = 30;
        var firstCell = sheet.Cells["C1"];
        firstCell.SetArrayFormula("=A1:A3*B1:B3", 3, 1);
        workbook.CalculateFormula();
        return workbook;
    }

    #endregion

    #region Non-Array Formula Cells

    [Fact]
    public void Execute_NonArrayFormulaCell_ReturnsNotFound()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetArrayFormulaResult>(res);

        Assert.False(result.IsArrayFormula);
        Assert.Contains("No array formula", result.Message);
    }

    [Fact]
    public void Execute_RegularFormulaCell_ReturnsNotFound()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Formula = "=A1*2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "B1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetArrayFormulaResult>(res);

        Assert.False(result.IsArrayFormula);
    }

    #endregion

    #region Array Formula Cells

    [Fact]
    public void Execute_ArrayFormulaCell_ReturnsIsArrayFormula()
    {
        var workbook = CreateWorkbookWithArrayFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetArrayFormulaResult>(res);

        Assert.True(result.IsArrayFormula);
    }

    [Fact]
    public void Execute_ArrayFormulaCell_ReturnsFormula()
    {
        var workbook = CreateWorkbookWithArrayFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetArrayFormulaResult>(res);

        Assert.NotEmpty(result.Formula ?? "");
    }

    [Fact]
    public void Execute_ArrayFormulaCell_ReturnsArrayRange()
    {
        var workbook = CreateWorkbookWithArrayFormula();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetArrayFormulaResult>(res);

        Assert.NotNull(result.ArrayRange);
    }

    #endregion
}
