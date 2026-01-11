using AsposeMcpServer.Handlers.Excel.GetCellAddress;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1", result);
        Assert.Contains("Row 0", result);
        Assert.Contains("Column 0", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B2", result);
        Assert.Contains("Row 1", result);
        Assert.Contains("Column 1", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("AA100", result);
        Assert.Contains("Row 99", result);
        Assert.Contains("Column 26", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Z1", result);
        Assert.Contains("Row 0", result);
        Assert.Contains("Column 25", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("XFD1", result);
        Assert.Contains("Row 0", result);
        Assert.Contains("Column 16383", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1048576", result);
        Assert.Contains("Row 1048575", result);
        Assert.Contains("Column 0", result);
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
