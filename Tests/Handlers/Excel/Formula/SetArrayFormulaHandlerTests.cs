using AsposeMcpServer.Handlers.Excel.Formula;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("B1:B3", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("B1:B2", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("B1:B2", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("B1:B2", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formula", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Range Size Tests

    [Fact]
    public void Execute_WithLargeRange_SetsFormula()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { 1, 2 },
            { 3, 4 },
            { 5, 6 },
            { 7, 8 },
            { 9, 10 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "C1:D5" },
            { "formula", "=A1:B5*2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("C1:D5", result.Message);
        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("C1:C3", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formula", result.Message, StringComparison.OrdinalIgnoreCase);
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

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A3" },
            { "formula", "=B1:B3*2" },
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
