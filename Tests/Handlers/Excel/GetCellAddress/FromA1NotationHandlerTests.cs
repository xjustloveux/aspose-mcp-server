using AsposeMcpServer.Handlers.Excel.GetCellAddress;
using AsposeMcpServer.Results.Excel.CellAddress;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.GetCellAddress;

public class FromA1NotationHandlerTests : ExcelHandlerTestBase
{
    private readonly FromA1NotationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_FromA1()
    {
        Assert.Equal("from_a1", _handler.Operation);
    }

    #endregion

    #region Basic A1 Notation Conversion

    [Fact]
    public void Execute_ConvertsA1ToIndices()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cellAddress", "A1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("A1", result.A1Notation);
        Assert.Equal(0, result.Row);
        Assert.Equal(0, result.Column);
    }

    [Fact]
    public void Execute_ConvertsB2ToIndices()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cellAddress", "B2" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("B2", result.A1Notation);
        Assert.Equal(1, result.Row);
        Assert.Equal(1, result.Column);
    }

    [Fact]
    public void Execute_ConvertsMultiLetterColumn()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cellAddress", "AA100" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("AA100", result.A1Notation);
        Assert.Equal(99, result.Row);
        Assert.Equal(26, result.Column);
    }

    [Fact]
    public void Execute_ConvertsZ1()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cellAddress", "Z1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("Z1", result.A1Notation);
        Assert.Equal(0, result.Row);
        Assert.Equal(25, result.Column);
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void Execute_WithLastColumn()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cellAddress", "XFD1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("XFD1", result.A1Notation);
        Assert.Equal(0, result.Row);
        Assert.Equal(16383, result.Column);
    }

    [Fact]
    public void Execute_WithLastRow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cellAddress", "A1048576" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("A1048576", result.A1Notation);
        Assert.Equal(1048575, result.Row);
        Assert.Equal(0, result.Column);
    }

    #endregion
}
