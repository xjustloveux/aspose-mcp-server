using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.RowColumn;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.RowColumn;

public class DeleteCellsHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteCellsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteCells()
    {
        Assert.Equal("delete_cells", _handler.Operation);
    }

    #endregion

    #region Various Ranges

    [Theory]
    [InlineData("A1")]
    [InlineData("A1:B2")]
    [InlineData("A1:D5")]
    public void Execute_WithVariousRanges_DeletesCells(string range)
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", range },
            { "shiftDirection", "up" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesCells()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "shiftDirection", "up" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        Assert.Null(GetCellValue(workbook, 0, 0));
        Assert.Null(GetCellValue(workbook, 0, 1));
        Assert.Null(GetCellValue(workbook, 1, 0));
        Assert.Null(GetCellValue(workbook, 1, 1));
    }

    [Fact]
    public void Execute_ReturnsRange()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "C3:D4" },
            { "shiftDirection", "left" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        AssertCellValue(workbook, 0, 0, "Data 1");
        AssertCellValue(workbook, 0, 1, "Data 2");
        AssertCellValue(workbook, 1, 0, "Data 3");
        AssertCellValue(workbook, 1, 1, "Data 4");
    }

    [Fact]
    public void Execute_ReturnsShiftDirection()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "shiftDirection", "left" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(GetCellValue(workbook, 0, 0));
        Assert.Null(GetCellValue(workbook, 0, 1));
        Assert.Null(GetCellValue(workbook, 1, 0));
        Assert.Null(GetCellValue(workbook, 1, 1));
    }

    #endregion

    #region Shift Direction Parameter

    [Fact]
    public void Execute_WithShiftUp_ShiftsCellsUp()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A2" },
            { "shiftDirection", "up" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(GetCellValue(workbook, 0, 0));
        Assert.Null(GetCellValue(workbook, 1, 0));
        AssertCellValue(workbook, 0, 1, "Data 2");
        AssertCellValue(workbook, 1, 1, "Data 4");
    }

    [Fact]
    public void Execute_WithShiftLeft_ShiftsCellsLeft()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:A2" },
            { "shiftDirection", "left" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        AssertCellValue(workbook, 0, 0, "Data 2");
        Assert.Null(GetCellValue(workbook, 0, 1));
        AssertCellValue(workbook, 1, 0, "Data 4");
        Assert.Null(GetCellValue(workbook, 1, 1));
    }

    [Fact]
    public void Execute_WithUpperCaseDirection_Works()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "shiftDirection", "LEFT" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(GetCellValue(workbook, 0, 0));
        Assert.Null(GetCellValue(workbook, 0, 1));
    }

    #endregion

    #region Sheet Index Parameter

    [Fact]
    public void Execute_WithSheetIndex_DeletesFromCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        workbook.Worksheets[1].Cells["A1"].PutValue("SheetData1");
        workbook.Worksheets[1].Cells["A2"].PutValue("SheetData2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "range", "A1:B2" },
            { "shiftDirection", "up" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(workbook.Worksheets[1].Cells[0, 0].Value);
        Assert.Null(workbook.Worksheets[1].Cells[1, 0].Value);
    }

    [Fact]
    public void Execute_DefaultSheetIndex_UsesFirstSheet()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "shiftDirection", "up" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(GetCellValue(workbook, 0, 0));
        Assert.Null(GetCellValue(workbook, 0, 1));
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shiftDirection", "up" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message);
    }

    [Fact]
    public void Execute_WithoutShiftDirection_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shiftDirection", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 },
            { "range", "A1:B2" },
            { "shiftDirection", "up" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Sheet index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", -1 },
            { "range", "A1:B2" },
            { "shiftDirection", "up" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Sheet index", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidRange_ThrowsException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "InvalidRange" },
            { "shiftDirection", "up" }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidShiftDirection_DefaultsToUp()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "shiftDirection", "invalid" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        Assert.Null(GetCellValue(workbook, 0, 0));
        Assert.Null(GetCellValue(workbook, 0, 1));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Data 1");
        sheet.Cells["B1"].PutValue("Data 2");
        sheet.Cells["A2"].PutValue("Data 3");
        sheet.Cells["B2"].PutValue("Data 4");
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
