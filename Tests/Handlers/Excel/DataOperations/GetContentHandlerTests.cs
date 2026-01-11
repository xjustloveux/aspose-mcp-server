using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataOperations;

public class GetContentHandlerTests : ExcelHandlerTestBase
{
    private readonly GetContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetContent()
    {
        Assert.Equal("get_content", _handler.Operation);
    }

    #endregion

    #region Basic Get Content Operations

    [Fact]
    public void Execute_ReturnsWorksheetContent()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Name");
        workbook.Worksheets[0].Cells["B1"].PutValue("Value");
        workbook.Worksheets[0].Cells["A2"].PutValue("Test");
        workbook.Worksheets[0].Cells["B2"].PutValue(100);
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Name", result);
        Assert.Contains("Value", result);
        Assert.Contains("Test", result);
        Assert.Contains("100", result);
    }

    [Fact]
    public void Execute_WithRange_ReturnsRangeContent()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("A1");
        workbook.Worksheets[0].Cells["B1"].PutValue("B1");
        workbook.Worksheets[0].Cells["A2"].PutValue("A2");
        workbook.Worksheets[0].Cells["B2"].PutValue("B2");
        workbook.Worksheets[0].Cells["C3"].PutValue("Outside");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("A1", result);
        Assert.Contains("B2", result);
        Assert.DoesNotContain("Outside", result);
    }

    [Fact]
    public void Execute_WithSheetIndex_GetsContentFromSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        workbook.Worksheets[0].Cells["A1"].PutValue("Sheet1Data");
        workbook.Worksheets[1].Cells["A1"].PutValue("Sheet2Data");
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Sheet2Data", result);
        Assert.DoesNotContain("Sheet1Data", result);
    }

    [Fact]
    public void Execute_WithMinimalData_ReturnsContent()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("OnlyCell");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("OnlyCell", result);
    }

    #endregion
}
