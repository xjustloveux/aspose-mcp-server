using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Handlers.Excel.PivotTable;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.PivotTable;

public class DeleteFieldExcelPivotTableHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteFieldExcelPivotTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteField()
    {
        Assert.Equal("delete_field", _handler.Operation);
    }

    #endregion

    #region Basic Delete Field Operations

    [Fact]
    public void Execute_DeletesFieldFromPivotTable()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "fieldName", "Category" },
            { "fieldType", "row" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithPivotTableAndFields()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(200);

        var pivotTables = worksheet.PivotTables;
        var idx = pivotTables.Add($"={worksheet.Name}!A1:B3", "D1", "TestPivot");
        var pivotTable = pivotTables[idx];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
        pivotTable.CalculateData();

        return workbook;
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPivotTableIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fieldName", "Category" },
            { "fieldType", "row" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutFieldName_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "fieldType", "row" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidFieldName_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "fieldName", "NonExistentField" },
            { "fieldType", "row" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
