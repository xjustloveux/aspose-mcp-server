using AsposeMcpServer.Handlers.Excel.FreezePanes;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.FreezePanes;

public class FreezeExcelPanesHandlerTests : ExcelHandlerTestBase
{
    private readonly FreezeExcelPanesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Freeze()
    {
        Assert.Equal("freeze", _handler.Operation);
    }

    #endregion

    #region Basic Freeze Operations

    [Fact]
    public void Execute_FreezesPanes()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 1 },
            { "column", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("frozen", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_FreezesOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "row", 2 },
            { "column", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("frozen", result.ToLower());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRow_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "column", 1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutColumn_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "row", 1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
