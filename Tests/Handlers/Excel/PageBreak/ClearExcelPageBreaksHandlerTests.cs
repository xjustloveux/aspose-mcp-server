using AsposeMcpServer.Handlers.Excel.PageBreak;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PageBreak;

public class ClearExcelPageBreaksHandlerTests : ExcelHandlerTestBase
{
    private readonly ClearExcelPageBreaksHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeClear()
    {
        Assert.Equal("clear", _handler.Operation);
    }

    #endregion

    #region Basic Clear Operations

    [Fact]
    public void Execute_ClearAll_ShouldRemoveAllPageBreaks()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.HorizontalPageBreaks.Add(10);
        worksheet.VerticalPageBreaks.Add(5);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "breakType", "all" }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(worksheet.HorizontalPageBreaks);
        Assert.Empty(worksheet.VerticalPageBreaks);
    }

    [Fact]
    public void Execute_ClearHorizontal_ShouldRemoveOnlyHorizontal()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.HorizontalPageBreaks.Add(10);
        worksheet.VerticalPageBreaks.Add(5);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "breakType", "horizontal" }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(worksheet.HorizontalPageBreaks);
        Assert.Single(worksheet.VerticalPageBreaks);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.HorizontalPageBreaks.Add(10);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "breakType", "all" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    #endregion
}
