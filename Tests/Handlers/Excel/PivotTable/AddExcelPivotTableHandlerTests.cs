using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.PivotTable;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.PivotTable;

public class AddExcelPivotTableHandlerTests : ExcelHandlerTestBase
{
    private readonly AddExcelPivotTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsPivotTable()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1:B5" },
            { "destCell", "D1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Single(workbook.Worksheets[0].PivotTables);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithName_UsesCustomName()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1:B5" },
            { "destCell", "D1" },
            { "name", "CustomPivot" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("CustomPivot", result);
    }

    [Fact]
    public void Execute_WithSheetIndex_AddsToSpecificSheet()
    {
        var workbook = CreateWorkbookWithData();
        workbook.Worksheets.Add("Sheet2");
        SetupDataOnSheet(workbook.Worksheets[1]);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "sourceRange", "A1:B5" },
            { "destCell", "D1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result.ToLower());
        Assert.Single(workbook.Worksheets[1].PivotTables);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSourceRange_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "destCell", "D1" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutDestCell_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithData();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sourceRange", "A1:B5" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithData()
    {
        var workbook = new Workbook();
        SetupDataOnSheet(workbook.Worksheets[0]);
        return workbook;
    }

    private static void SetupDataOnSheet(Worksheet worksheet)
    {
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(150);
        worksheet.Cells["A5"].PutValue("B");
        worksheet.Cells["B5"].PutValue(250);
    }

    #endregion
}
