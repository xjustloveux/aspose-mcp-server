using AsposeMcpServer.Handlers.Excel.PageBreak;
using AsposeMcpServer.Results.Excel.PageBreak;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PageBreak;

public class GetExcelPageBreaksHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelPageBreaksHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeGet()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoPageBreaks_ShouldReturnEmpty()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPageBreaksExcelResult>(res);
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void Execute_WithPageBreaks_ShouldReturnInfo()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.HorizontalPageBreaks.Add(10);
        worksheet.VerticalPageBreaks.Add(5);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetPageBreaksExcelResult>(res);
        Assert.Equal(2, result.Count);
        Assert.Equal(2, result.Items.Count);
    }

    #endregion
}
