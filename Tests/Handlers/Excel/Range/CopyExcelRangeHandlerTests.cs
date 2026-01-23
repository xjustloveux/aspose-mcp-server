using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Range;

public class CopyExcelRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly CopyExcelRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Copy()
    {
        Assert.Equal("copy", _handler.Operation);
    }

    #endregion

    #region Preserve Source

    [Fact]
    public void Execute_PreservesSourceRange()
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

        Assert.Equal("Original", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("Original", workbook.Worksheets[0].Cells["B1"].Value);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("A1:B2", result.Message);
        Assert.Contains("D1", result.Message);
        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Copy Operations

    [Fact]
    public void Execute_CopiesRangeToDestination()
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("A", workbook.Worksheets[0].Cells["D1"].Value);
        Assert.Equal("B", workbook.Worksheets[0].Cells["E1"].Value);
        Assert.Equal("C", workbook.Worksheets[0].Cells["D2"].Value);
        Assert.Equal("D", workbook.Worksheets[0].Cells["E2"].Value);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1", "C1")]
    [InlineData("A1:B1", "D1")]
    [InlineData("A1:A2", "C1")]
    public void Execute_CopiesVariousRanges(string sourceRange, string destCell)
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "1", "2", "" },
            { "3", "4", "" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", sourceRange },
            { "destCell", destCell }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("copied", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Cross-Sheet Copy

    [Fact]
    public void Execute_WithDestSheetIndex_CopiesToDifferentSheet()
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

        Assert.Equal("Source", workbook.Worksheets[0].Cells["A1"].Value);
        Assert.Equal("Source", workbook.Worksheets[1].Cells["A1"].Value);
    }

    [Fact]
    public void Execute_WithSourceSheetIndex_CopiesFromDifferentSheet()
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
    }

    #endregion

    #region Copy Options

    [Fact]
    public void Execute_WithCopyOptionsAll_CopiesValueAndFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1" },
            { "destCell", "B1" },
            { "copyOptions", "All" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Test", workbook.Worksheets[0].Cells["B1"].Value);
        Assert.True(workbook.Worksheets[0].Cells["B1"].GetStyle().Font.IsBold);
    }

    [Fact]
    public void Execute_DefaultCopyOptions_CopiesAll()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1" },
            { "destCell", "B1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Test", workbook.Worksheets[0].Cells["B1"].Value);
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
