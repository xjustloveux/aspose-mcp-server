using AsposeMcpServer.Handlers.Excel.GetCellAddress;
using AsposeMcpServer.Results.Excel.CellAddress;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.GetCellAddress;

public class FromIndexHandlerTests : ExcelHandlerTestBase
{
    private readonly FromIndexHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_FromIndex()
    {
        Assert.Equal("from_index", _handler.Operation);
    }

    #endregion

    #region Basic Index Conversion

    [Fact]
    public void Execute_ConvertsRow0Col0ToA1()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 0 },
            { "column", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("A1", result.A1Notation);
        Assert.Equal(0, result.Row);
        Assert.Equal(0, result.Column);
    }

    [Fact]
    public void Execute_ConvertsRow1Col1ToB2()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 1 },
            { "column", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("B2", result.A1Notation);
        Assert.Equal(1, result.Row);
        Assert.Equal(1, result.Column);
    }

    [Fact]
    public void Execute_ConvertsRow99Col26ToAA100()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 99 },
            { "column", 26 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("AA100", result.A1Notation);
        Assert.Equal(99, result.Row);
        Assert.Equal(26, result.Column);
    }

    [Fact]
    public void Execute_ConvertsCol25ToZ()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 0 },
            { "column", 25 }
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
            { "row", 0 },
            { "column", 16383 }
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
            { "row", 1048575 },
            { "column", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<CellAddressResult>(res);

        Assert.Equal("A1048576", result.A1Notation);
        Assert.Equal(1048575, result.Row);
        Assert.Equal(0, result.Column);
    }

    [Fact]
    public void Execute_WithNegativeRow_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", -1 },
            { "column", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeColumn_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 0 },
            { "column", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithRowOutOfRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 1048576 },
            { "column", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithColumnOutOfRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 0 },
            { "column", 16384 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
