using AsposeMcpServer.Handlers.Excel.PageBreak;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PageBreak;

public class RemoveExcelPageBreakHandlerTests : ExcelHandlerTestBase
{
    private readonly RemoveExcelPageBreakHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeRemove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    #endregion

    #region Basic Remove Operations

    [Fact]
    public void Execute_ShouldRemoveHorizontalPageBreak()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.HorizontalPageBreaks.Add(10);
        Assert.Single(worksheet.HorizontalPageBreaks);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "breakType", "horizontal" },
            { "breakIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(worksheet.HorizontalPageBreaks);
    }

    [Fact]
    public void Execute_ShouldRemoveVerticalPageBreak()
    {
        var workbook = CreateEmptyWorkbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.VerticalPageBreaks.Add(5);
        Assert.Single(worksheet.VerticalPageBreaks);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "breakType", "vertical" },
            { "breakIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Empty(worksheet.VerticalPageBreaks);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingBreakType_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "breakIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("breakType", ex.Message);
    }

    [Fact]
    public void Execute_WithMissingBreakIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "breakType", "horizontal" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("breakIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidBreakIndex_ShouldThrow()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "breakType", "horizontal" },
            { "breakIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
