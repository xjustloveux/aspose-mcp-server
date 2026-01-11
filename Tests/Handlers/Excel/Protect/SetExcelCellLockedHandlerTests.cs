using AsposeMcpServer.Handlers.Excel.Protect;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Protect;

public class SetExcelCellLockedHandlerTests : ExcelHandlerTestBase
{
    private readonly SetExcelCellLockedHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetCellLocked()
    {
        Assert.Equal("set_cell_locked", _handler.Operation);
    }

    #endregion

    #region Multiple Sheets

    [Fact]
    public void Execute_WithSheetIndex_ModifiesCorrectSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "sheetIndex", 1 },
            { "locked", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 1", result.ToLower());
        Assert.True(workbook.Worksheets[1].Cells["A1"].GetStyle().IsLocked);
    }

    #endregion

    #region Basic Lock Operations

    [Fact]
    public void Execute_LocksCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "locked", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("locked", result.ToLower());
        Assert.Contains("A1:B2", result);
        Assert.True(workbook.Worksheets[0].Cells["A1"].GetStyle().IsLocked);
        AssertModified(context);
    }

    [Fact]
    public void Execute_UnlocksCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "locked", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("unlocked", result.ToLower());
        Assert.False(workbook.Worksheets[0].Cells["A1"].GetStyle().IsLocked);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DefaultsToUnlocked()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("unlocked", result.ToLower());
    }

    [Fact]
    public void Execute_WithSingleCell_LocksSingleCell()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "C5" },
            { "locked", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("C5", result);
        Assert.True(workbook.Worksheets[0].Cells["C5"].GetStyle().IsLocked);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithEmptyRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
