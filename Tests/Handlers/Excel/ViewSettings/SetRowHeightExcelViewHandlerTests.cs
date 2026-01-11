using AsposeMcpServer.Handlers.Excel.ViewSettings;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.ViewSettings;

public class SetRowHeightExcelViewHandlerTests : ExcelHandlerTestBase
{
    private readonly SetRowHeightExcelViewHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetRowHeight()
    {
        Assert.Equal("set_row_height", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsRowHeight()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "height", 25.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("height set to 25", result.ToLower());
        Assert.Equal(25.0, workbook.Worksheets[0].Cells.GetRowHeight(0));
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_SetsRowHeightOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "rowIndex", 3 },
            { "height", 30.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("row 3", result.ToLower());
        Assert.Equal(30.0, workbook.Worksheets[1].Cells.GetRowHeight(3));
    }

    #endregion

    #region Boundary Condition Tests

    [Fact]
    public void Execute_WithoutRowIndex_UsesDefaultRowIndex()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "height", 25.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("row 0", result.ToLower());
        Assert.Contains("height set to 25", result.ToLower());
    }

    [Fact]
    public void Execute_WithoutHeight_UsesDefaultHeight()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("height set to 15", result.ToLower());
    }

    [Fact]
    public void Execute_WithNegativeRowIndex_HandlesGracefully()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", -1 },
            { "height", 25.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("height set to", result.ToLower());
    }

    [Fact]
    public void Execute_WithZeroHeight_SetsZeroHeight()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "height", 0.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("height set to", result.ToLower());
    }

    [Fact]
    public void Execute_WithNegativeHeight_ThrowsCellsException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "height", -10.0 }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithSmallPositiveHeight_SetsHeight()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "height", 0.1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("height set to", result.ToLower());
    }

    [Fact]
    public void Execute_WithLargeHeight_SetsHeight()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "height", 400.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("height set to 400", result.ToLower());
    }

    [Fact]
    public void Execute_WithHeightExceedingMax_ThrowsException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "height", 500.0 }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithRowIndexZero_SetsFirstRowHeight()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "height", 20.0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("row 0", result.ToLower());
    }

    #endregion
}
