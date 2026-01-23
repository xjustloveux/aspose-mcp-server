using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Filter;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Filter;

public class RemoveExcelFilterHandlerTests : ExcelHandlerTestBase
{
    private readonly RemoveExcelFilterHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Remove()
    {
        Assert.Equal("remove", _handler.Operation);
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_RemovesFromCorrectSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[0].Cells["A1"].Value = "Test";
        workbook.Worksheets[0].AutoFilter.Range = "A1:B10";
        workbook.Worksheets[1].Cells["A1"].Value = "Test";
        workbook.Worksheets[1].AutoFilter.Range = "A1:C5";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("A1:B10", workbook.Worksheets[0].AutoFilter.Range);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithFilter();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithFilter()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["A2"].Value = "Item1";
        sheet.AutoFilter.Range = "A1:A10";
        return workbook;
    }

    #endregion

    #region Basic Remove Operations

    [Fact]
    public void Execute_RemovesAutoFilter()
    {
        var workbook = CreateWorkbookWithFilter();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Auto filter removed", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSheetIndex()
    {
        var workbook = CreateWorkbookWithFilter();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("sheet 0", result.Message);
    }

    [Fact]
    public void Execute_CanBeCalledWithoutExistingFilter()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Auto filter removed", result.Message);
    }

    #endregion
}
