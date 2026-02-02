using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.RowColumn;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.RowColumn;

public class InsertColumnHandlerTests : ExcelHandlerTestBase
{
    private readonly InsertColumnHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_InsertColumn()
    {
        Assert.Equal("insert_column", _handler.Operation);
    }

    #endregion

    #region Count Parameter

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    public void Execute_WithVariousCounts_InsertsColumns(int count)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", count }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        AssertCellValue(workbook, 0, count, "Col A");
        AssertCellValue(workbook, 0, count + 1, "Col B");
        AssertCellValue(workbook, 0, count + 2, "Col C");
        for (var i = 0; i < count; i++)
            Assert.Null(GetCellValue(workbook, 0, i));
    }

    #endregion

    #region Basic Insert Operations

    [Fact]
    public void Execute_InsertsColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        AssertCellValue(workbook, 0, 0, "Col A");
        Assert.Null(GetCellValue(workbook, 0, 1));
        AssertCellValue(workbook, 0, 2, "Col B");
        AssertCellValue(workbook, 0, 3, "Col C");
    }

    [Fact]
    public void Execute_ReturnsColumnIndex()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        AssertCellValue(workbook, 0, 0, "Col A");
        AssertCellValue(workbook, 0, 1, "Col B");
        Assert.Null(GetCellValue(workbook, 0, 2));
        AssertCellValue(workbook, 0, 3, "Col C");
    }

    [Fact]
    public void Execute_ReturnsCount()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", 3 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        for (var i = 0; i < 3; i++)
            Assert.Null(GetCellValue(workbook, 0, i));
        AssertCellValue(workbook, 0, 3, "Col A");
        AssertCellValue(workbook, 0, 4, "Col B");
        AssertCellValue(workbook, 0, 5, "Col C");
    }

    [Fact]
    public void Execute_DefaultCount_InsertsOneColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(GetCellValue(workbook, 0, 0));
        AssertCellValue(workbook, 0, 1, "Col A");
        AssertCellValue(workbook, 0, 2, "Col B");
        AssertCellValue(workbook, 0, 3, "Col C");
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_InsertsOnCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        workbook.Worksheets[1].Cells["A1"].PutValue("Sheet2Col");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(workbook.Worksheets[1].Cells[0, 0].Value);
        Assert.Equal("Sheet2Col", workbook.Worksheets[1].Cells[0, 1].Value);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_UsesFirstSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        workbook.Worksheets[0].Cells["A1"].PutValue("FirstSheetCol");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(workbook.Worksheets[0].Cells[0, 0].Value);
        Assert.Equal("FirstSheetCol", workbook.Worksheets[0].Cells[0, 1].Value);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Sheet index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeColumnIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Invalid column index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", -1 },
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Sheet index", ex.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Execute_WithZeroOrNegativeCount_ThrowsArgumentException(int count)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", count }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Count must be greater than 0", ex.Message);
    }

    [Fact]
    public void Execute_WithColumnIndexZero_InsertsAtFirstColumn()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(GetCellValue(workbook, 0, 0));
        AssertCellValue(workbook, 0, 1, "Col A");
    }

    [Fact]
    public void Execute_WithLargeCount_InsertsMultipleColumns()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "count", 100 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        AssertCellValue(workbook, 0, 100, "Col A");
        AssertCellValue(workbook, 0, 101, "Col B");
        AssertCellValue(workbook, 0, 102, "Col C");
        Assert.Null(GetCellValue(workbook, 0, 0));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Col A");
        sheet.Cells["B1"].PutValue("Col B");
        sheet.Cells["C1"].PutValue("Col C");
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
