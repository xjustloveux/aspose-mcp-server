using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Formula;

public class AddFormulaHandlerTests : ExcelHandlerTestBase
{
    private readonly AddFormulaHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_AddsFormulaToCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = 100;
        workbook.Worksheets[1].Cells["B1"].Value = 200;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "formula", "=A1+B1" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("=A1+B1", workbook.Worksheets[1].Cells["C1"].Formula);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsFormula()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 10, 20 },
            { 30, 40 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "formula", "=A1+B1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Formula added", result);
        Assert.Contains("C1", result);
        AssertModified(context);
    }

    [Theory]
    [InlineData("=SUM(A1:B1)")]
    [InlineData("=AVERAGE(A1:B2)")]
    [InlineData("=MAX(A1,B1,A2,B2)")]
    public void Execute_AddsVariousFormulas(string formula)
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 10, 20 },
            { 30, 40 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "formula", formula }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Formula added", result);
        Assert.Contains(formula, result);
    }

    [Fact]
    public void Execute_SetsFormulaOnCell()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 10, 20 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "formula", "=A1+B1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("=A1+B1", workbook.Worksheets[0].Cells["C1"].Formula);
    }

    #endregion

    #region AutoCalculate

    [Fact]
    public void Execute_WithAutoCalculate_CalculatesFormula()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 10, 20 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "formula", "=A1+B1" },
            { "autoCalculate", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(30.0, workbook.Worksheets[0].Cells["C1"].Value);
    }

    [Fact]
    public void Execute_WithAutoCalculateFalse_DoesNotCalculate()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 10, 20 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" },
            { "formula", "=A1+B1" },
            { "autoCalculate", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("=A1+B1", workbook.Worksheets[0].Cells["C1"].Formula);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formula", "=SUM(A1:A10)" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("cell", ex.Message);
    }

    [Fact]
    public void Execute_WithoutFormula_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "C1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("formula", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidFormula_ReturnsWarning()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "formula", "=INVALIDFUNCTION()" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#NAME?", result);
    }

    #endregion
}
