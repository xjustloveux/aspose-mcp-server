using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.RowColumn;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.RowColumn;

public class InsertRowHandlerTests : ExcelHandlerTestBase
{
    private readonly InsertRowHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertRow()
    {
        Assert.Equal("insert_row", _handler.Operation);
    }

    #endregion

    #region Count Parameter

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(10)]
    public void Execute_WithVariousCounts_InsertsRows(int count)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "count", count }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"{count} row(s)", result);
    }

    #endregion

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsRow()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Inserted", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsRowIndex()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("row 2", result);
    }

    [Fact]
    public void Execute_ReturnsCount()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "count", 3 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("3 row(s)", result);
    }

    [Fact]
    public void Execute_DefaultCount_InsertsOneRow()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1 row(s)", result);
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_InsertsOnCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "rowIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Inserted", result);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_UsesFirstSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Inserted", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRowIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rowIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "rowIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Sheet index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeRowIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Invalid row index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", -1 },
            { "rowIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Sheet index", ex.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Execute_WithZeroOrNegativeCount_AcceptsValue(int count)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "count", count }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains($"{count} row(s)", result);
    }

    [Fact]
    public void Execute_WithRowIndexZero_InsertsAtFirstRow()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Inserted", result);
        Assert.Contains("row 0", result);
    }

    [Fact]
    public void Execute_WithLargeCount_InsertsMultipleRows()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "count", 100 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("100 row(s)", result);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Data 1");
        sheet.Cells["A2"].PutValue("Data 2");
        sheet.Cells["A3"].PutValue("Data 3");
        return workbook;
    }

    private new static Workbook CreateWorkbookWithSheets(int sheetCount)
    {
        var workbook = new Workbook();
        for (var i = 1; i < sheetCount; i++)
            workbook.Worksheets.Add();
        return workbook;
    }

    #endregion
}
