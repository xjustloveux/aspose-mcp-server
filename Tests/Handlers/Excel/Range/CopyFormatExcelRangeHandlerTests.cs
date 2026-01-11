using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Range;

public class CopyFormatExcelRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly CopyFormatExcelRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_CopyFormat()
    {
        Assert.Equal("copy_format", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_CopiesFormatOnSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = "Source";
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[1].Cells["A1"].SetStyle(style);
        workbook.Worksheets[1].Cells["B1"].Value = "Dest";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "destTarget", "B1" },
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.True(workbook.Worksheets[1].Cells["B1"].GetStyle().Font.IsBold);
    }

    #endregion

    #region Preserve Source

    [Fact]
    public void Execute_PreservesSourceFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Source" } });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "destTarget", "B1" }
        });

        _handler.Execute(context, parameters);

        Assert.True(workbook.Worksheets[0].Cells["A1"].GetStyle().Font.IsBold);
    }

    #endregion

    #region Basic Copy Format Operations

    [Fact]
    public void Execute_CopiesFormatToDestination()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Source" } });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        style.Font.Size = 14;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        workbook.Worksheets[0].Cells["B1"].Value = "Dest";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "destTarget", "B1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("copied", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(workbook.Worksheets[0].Cells["B1"].GetStyle().Font.IsBold);
        Assert.Equal(14, workbook.Worksheets[0].Cells["B1"].GetStyle().Font.Size);
        AssertModified(context);
    }

    [Theory]
    [InlineData("A1", "B1")]
    [InlineData("A1:B2", "D1")]
    [InlineData("A1", "C3")]
    public void Execute_CopiesFormatToVariousDestinations(string range, string destTarget)
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "1", "2", "", "" },
            { "3", "4", "", "" },
            { "", "", "", "" }
        });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", range },
            { "destTarget", destTarget }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("copied", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Copy Value Option

    [Fact]
    public void Execute_WithCopyValueFalse_CopiesOnlyFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Source" } });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        workbook.Worksheets[0].Cells["B1"].Value = "OriginalDest";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "destTarget", "B1" },
            { "copyValue", false }
        });

        _handler.Execute(context, parameters);

        Assert.True(workbook.Worksheets[0].Cells["B1"].GetStyle().Font.IsBold);
        Assert.Equal("OriginalDest", workbook.Worksheets[0].Cells["B1"].Value);
    }

    [Fact]
    public void Execute_WithCopyValueTrue_CopiesFormatAndValue()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Source" } });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        workbook.Worksheets[0].Cells["B1"].Value = "OriginalDest";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "destTarget", "B1" },
            { "copyValue", true }
        });

        _handler.Execute(context, parameters);

        Assert.True(workbook.Worksheets[0].Cells["B1"].GetStyle().Font.IsBold);
        Assert.Equal("Source", workbook.Worksheets[0].Cells["B1"].Value);
    }

    [Fact]
    public void Execute_DefaultCopyValue_CopiesOnlyFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Source" } });
        var style = workbook.CreateStyle();
        style.Font.IsBold = true;
        workbook.Worksheets[0].Cells["A1"].SetStyle(style);
        workbook.Worksheets[0].Cells["B1"].Value = "KeepThis";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "destTarget", "B1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("KeepThis", workbook.Worksheets[0].Cells["B1"].Value);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "destTarget", "B1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutDestTarget_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("dest", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_WithCopyValueFalse_ReturnsFormatCopiedMessage()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "destTarget", "B1" },
            { "copyValue", false }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Format copied", result);
    }

    [Fact]
    public void Execute_WithCopyValueTrue_ReturnsFormatWithValuesCopiedMessage()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "destTarget", "B1" },
            { "copyValue", true }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Format with values copied", result);
    }

    #endregion
}
