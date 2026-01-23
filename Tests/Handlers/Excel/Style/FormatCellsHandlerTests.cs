using AsposeMcpServer.Handlers.Excel.Style;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Style;

public class FormatCellsHandlerTests : ExcelHandlerTestBase
{
    private readonly FormatCellsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Format()
    {
        Assert.Equal("format", _handler.Operation);
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
            { "bold", true }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Vertical Alignment

    [Theory]
    [InlineData("top")]
    [InlineData("center")]
    [InlineData("bottom")]
    public void Execute_WithVerticalAlignment_SetsVerticalAlignment(string alignment)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "verticalAlignment", alignment }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Horizontal Alignment

    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    public void Execute_WithHorizontalAlignmentVariations_SetsAlignment(string alignment)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "horizontalAlignment", alignment }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Multiple Ranges

    [Fact]
    public void Execute_WithMultipleRanges_FormatsAllRanges()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "ranges", "[\"A1:B2\", \"C3:D4\"]" },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_FormatsCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "range", "A1" },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Format Operations

    [Fact]
    public void Execute_FormatsCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B5" },
            { "bold", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFontName_SetsFontName()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "fontName", "Arial" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithFontSize_SetsFontSize()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "fontSize", 14 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithBackgroundColor_SetsBackgroundColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "backgroundColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithBorderStyle_SetsBorder()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" },
            { "borderStyle", "thin" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithAlignment_SetsAlignment()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "horizontalAlignment", "center" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Pattern Types

    [Theory]
    [InlineData("solid")]
    [InlineData("gray50")]
    [InlineData("gray75")]
    [InlineData("gray25")]
    [InlineData("horizontalstripe")]
    [InlineData("verticalstripe")]
    [InlineData("diagonalstripe")]
    [InlineData("reversediagonalstripe")]
    [InlineData("diagonalcrosshatch")]
    [InlineData("thickdiagonalcrosshatch")]
    [InlineData("thinhorizontalstripe")]
    [InlineData("thinverticalstripe")]
    [InlineData("thinreversediagonalstripe")]
    [InlineData("thindiagonalstripe")]
    [InlineData("thinhorizontalcrosshatch")]
    [InlineData("thindiagonalcrosshatch")]
    public void Execute_WithPatternType_SetsPattern(string patternType)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "patternType", patternType },
            { "backgroundColor", "#FFFF00" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithPatternColor_SetsPatternColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "patternType", "gray50" },
            { "backgroundColor", "#FFFF00" },
            { "patternColor", "#0000FF" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Number Format

    [Fact]
    public void Execute_WithNumberFormat_SetsNumberFormat()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "numberFormat", "#,##0.00" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNumericFormatNumber_SetsFormatNumber()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "numberFormat", "4" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Border Styles

    [Theory]
    [InlineData("none")]
    [InlineData("thin")]
    [InlineData("medium")]
    [InlineData("thick")]
    [InlineData("dotted")]
    [InlineData("dashed")]
    [InlineData("double")]
    public void Execute_WithBorderStyleVariations_SetsBorder(string borderStyle)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "borderStyle", borderStyle }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithBorderColor_SetsBorderColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "borderStyle", "thin" },
            { "borderColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Font Formatting

    [Fact]
    public void Execute_WithItalic_SetsItalic()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "italic", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithFontColor_SetsFontColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "fontColor", "#0000FF" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("formatted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
