using AsposeMcpServer.Handlers.Excel.Style;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Style;

public class GetCellFormatHandlerTests : ExcelHandlerTestBase
{
    private readonly GetCellFormatHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetFormat()
    {
        Assert.Equal("get_format", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutCellOrRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsFormatForCell()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Test");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1", result);
        Assert.Contains("format", result.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_GetsFormatForRange()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result.ToLower());
    }

    [Fact]
    public void Execute_WithSpecificFields_ReturnsOnlyRequestedFields()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" },
            { "fields", "font" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("fontName", result);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cell", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
    }

    #endregion
}
