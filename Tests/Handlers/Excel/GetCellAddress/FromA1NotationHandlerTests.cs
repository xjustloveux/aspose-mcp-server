using AsposeMcpServer.Handlers.Excel.GetCellAddress;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1", result);
        Assert.Contains("Row 0", result);
        Assert.Contains("Column 0", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("B2", result);
        Assert.Contains("Row 1", result);
        Assert.Contains("Column 1", result);
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("AA100", result);
        Assert.Contains("Row 99", result);
        Assert.Contains("Column 26", result);
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
            { "cellAddress", "XFD1" }
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
            { "cellAddress", "A1048576" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1048576", result);
        Assert.Contains("Row 1048575", result);
        Assert.Contains("Column 0", result);
    }

    #endregion
}
