using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Results.Excel.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentResult>(res);

        Assert.NotNull(result.Rows);
        Assert.Equal(2, result.Rows.Count);
        Assert.Contains(result.Rows[0].Values, v => v?.ToString() == "Name");
        Assert.Contains(result.Rows[0].Values, v => v?.ToString() == "Value");
        Assert.Contains(result.Rows[1].Values, v => v?.ToString() == "Test");
        // ReSharper disable once CompareOfFloatsByEqualityOperator - Exact integer value 100 comparison is safe
        Assert.Contains(result.Rows[1].Values,
            v => v?.ToString() == "100" || v is (int or double) and 100);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentResult>(res);

        Assert.NotNull(result.Rows);
        Assert.Equal(2, result.Rows.Count);
        Assert.Contains(result.Rows[0].Values, v => v?.ToString() == "A1");
        Assert.Contains(result.Rows[1].Values, v => v?.ToString() == "B2");
        Assert.DoesNotContain(result.Rows, r => r.Values.Any(v => v?.ToString() == "Outside"));
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentResult>(res);

        Assert.NotNull(result.Rows);
        Assert.Contains(result.Rows, r => r.Values.Any(v => v?.ToString() == "Sheet2Data"));
        Assert.DoesNotContain(result.Rows, r => r.Values.Any(v => v?.ToString() == "Sheet1Data"));
    }

    [Fact]
    public void Execute_WithMinimalData_ReturnsContent()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("OnlyCell");
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetContentResult>(res);

        Assert.NotNull(result.Rows);
        Assert.Contains(result.Rows, r => r.Values.Any(v => v?.ToString() == "OnlyCell"));
    }

    #endregion
}
