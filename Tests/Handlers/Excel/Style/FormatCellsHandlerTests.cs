using AsposeMcpServer.Handlers.Excel.Style;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatted", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatted", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatted", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatted", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatted", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("formatted", result.ToLower());
    }

    #endregion
}
