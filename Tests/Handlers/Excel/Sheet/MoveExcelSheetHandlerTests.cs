using AsposeMcpServer.Handlers.Excel.Sheet;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Sheet;

public class MoveExcelSheetHandlerTests : ExcelHandlerTestBase
{
    private readonly MoveExcelSheetHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Move()
    {
        Assert.Equal("move", _handler.Operation);
    }

    #endregion

    #region Basic Move Operations

    [Theory]
    [InlineData(0, 1)]
    [InlineData(0, 2)]
    [InlineData(2, 0)]
    [InlineData(1, 0)]
    public void Execute_MovesSheetToVariousPositions(int from, int to)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        var sheetName = workbook.Worksheets[from].Name;
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", from },
            { "targetIndex", to }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(sheetName, workbook.Worksheets[to].Name);
    }

    #endregion

    #region InsertAt Parameter

    [Fact]
    public void Execute_WithInsertAt_MovesToPosition()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 2 },
            { "insertAt", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Sheet3", workbook.Worksheets[0].Name);
    }

    #endregion

    #region Same Position

    [Fact]
    public void Execute_SamePosition_ReturnsNoMoveNeeded()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "targetIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("no move needed", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsPositionsInMessage()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "targetIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0", result);
        Assert.Contains("2", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "targetIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sheetIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutTargetOrInsertAt_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("targetIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(10, 0)]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException(int invalidIndex, int targetIndex)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", invalidIndex },
            { "targetIndex", targetIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(0, -1)]
    [InlineData(0, 10)]
    public void Execute_WithInvalidTargetIndex_ThrowsArgumentException(int sheetIndex, int invalidTarget)
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", sheetIndex },
            { "targetIndex", invalidTarget }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
