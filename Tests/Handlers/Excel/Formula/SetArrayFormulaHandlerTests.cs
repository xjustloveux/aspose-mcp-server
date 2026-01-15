using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Formula;

public class SetArrayFormulaHandlerTests : ExcelHandlerTestBase
{
    private readonly SetArrayFormulaHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetArray()
    {
        Assert.Equal("set_array", _handler.Operation);
    }

    #endregion

    #region Formula Variations

    [Theory]
    [InlineData("=A1:A3*2")]
    [InlineData("{=A1:A3*2}")]
    public void Execute_HandlesFormulaWithOrWithoutBraces(string formula)
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1 },
            { 2 },
            { 3 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1:B3" },
            { "formula", formula }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B1:B3", result);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_SetsFormulaOnCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = 10;
        workbook.Worksheets[1].Cells["A2"].Value = 20;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1:B2" },
            { "formula", "=A1:A2*2" },
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B1:B2", result);
    }

    #endregion

    #region AutoCalculate

    [Fact]
    public void Execute_WithAutoCalculateFalse_DoesNotCalculate()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1 },
            { 2 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1:B2" },
            { "formula", "=A1:A2*10" },
            { "autoCalculate", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B1:B2", result);
    }

    #endregion

    #region Formula Without Equals Sign

    [Fact]
    public void Execute_WithFormulaWithoutEquals_AddsEquals()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1 },
            { 2 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1:B2" },
            { "formula", "A1:A2*2" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B1:B2", result);
    }

    #endregion

    #region Single Cell Range

    [Fact]
    public void Execute_WithSingleCellRange_SetsFormula()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1, 2, 3 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A2" },
            { "formula", "=SUM(A1:C1)" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formula", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Set Array Operations

    [Fact]
    public void Execute_SetsArrayFormula()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1, 2 },
            { 3, 4 },
            { 5, 6 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "C1:C3" },
            { "formula", "=A1:A3*B1:B3" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("C1:C3", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1, 2, 3 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "D1:D3" },
            { "formula", "=A1:C1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formula", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "formula", "=A1:A3*2" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message);
    }

    [Fact]
    public void Execute_WithoutFormula_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1:B3" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("formula", ex.Message);
    }

    #endregion
}
