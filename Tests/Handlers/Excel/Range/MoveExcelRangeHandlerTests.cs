using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Range;

public class MoveExcelRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly MoveExcelRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Move()
    {
        Assert.Equal("move", _handler.Operation);
    }

    #endregion

    #region Preserve Other Cells

    [Fact]
    public void Execute_PreservesOtherCells()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Move", "Keep1" },
            { "Keep2", "Keep3" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1" },
            { "destCell", "C1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Keep1", workbook.Worksheets[0].Cells["B1"].Value);
        Assert.Equal("Keep2", workbook.Worksheets[0].Cells["A2"].Value);
        Assert.Equal("Move", workbook.Worksheets[0].Cells["C1"].Value);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSuccessMessage()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1:B2" },
            { "destCell", "D1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1:B2", result);
        Assert.Contains("D1", result);
        Assert.Contains("moved", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Move Operations

    [Fact]
    public void Execute_MovesRangeToDestination()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A", "B" },
            { "C", "D" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1:B2" },
            { "destCell", "D1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("moved", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("A", workbook.Worksheets[0].Cells["D1"].Value);
        Assert.Equal("D", workbook.Worksheets[0].Cells["E2"].Value);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1", "C1")]
    [InlineData("A1:B1", "D1")]
    [InlineData("A1:A2", "C1")]
    public void Execute_MovesVariousRanges(string sourceRange, string destCell)
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "1", "2", "", "" },
            { "3", "4", "", "" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", sourceRange },
            { "destCell", destCell }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("moved", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Source Cleared

    [Fact]
    public void Execute_ClearsSourceRange()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Original" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1" },
            { "destCell", "B1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Original", workbook.Worksheets[0].Cells["B1"].Value);
    }

    [Fact]
    public void Execute_ClearsEntireSourceRange()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A", "B" },
            { "C", "D" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1:B2" },
            { "destCell", "D1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("", workbook.Worksheets[0].Cells["B1"].StringValue);
        Assert.Equal("", workbook.Worksheets[0].Cells["A2"].StringValue);
        Assert.Equal("", workbook.Worksheets[0].Cells["B2"].StringValue);
    }

    #endregion

    #region Cross-Sheet Move

    [Fact]
    public void Execute_WithDestSheetIndex_MovesToDifferentSheet()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Source" } });
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1" },
            { "destCell", "A1" },
            { "destSheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Source", workbook.Worksheets[1].Cells["A1"].Value);
    }

    [Fact]
    public void Execute_WithSourceSheetIndex_MovesFromDifferentSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = "FromSheet2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1" },
            { "destCell", "A1" },
            { "sourceSheetIndex", 1 },
            { "destSheetIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("FromSheet2", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("", workbook.Worksheets[1].Cells["A1"].StringValue);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSourceRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "destCell", "B1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sourceRange", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutDestCell_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("destCell", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("InvalidRange", "B1")]
    [InlineData("A1", "InvalidCell")]
    public void Execute_WithInvalidRangeOrCell_ThrowsException(string sourceRange, string destCell)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", sourceRange },
            { "destCell", destCell }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
