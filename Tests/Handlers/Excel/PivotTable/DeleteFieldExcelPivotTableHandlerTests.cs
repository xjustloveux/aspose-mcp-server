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

    [Fact]
    public void Execute_WithSheetIndex_UsesCorrectSheet()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 0 },
            { "pivotTableIndex", 0 },
            { "fieldName", "Category" },
            { "fieldType", "row" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeletesColumnField()
    {
        var workbook = CreateWorkbookWithPivotTableAndMultipleFieldTypes();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "fieldName", "Product" },
            { "fieldType", "column" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("column", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DeletesDataField()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "fieldName", "Value" },
            { "fieldType", "data" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFieldNotInArea_ReturnsSuccessWithMayAlreadyBeRemovedMessage()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "fieldName", "Value" },
            { "fieldType", "row" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.True(
            result.Message.Contains("removed", StringComparison.OrdinalIgnoreCase) ||
            result.Message.Contains("may already be removed", StringComparison.OrdinalIgnoreCase));
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

    private static Workbook CreateWorkbookWithPivotTableAndMultipleFieldTypes()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Region");
        worksheet.Cells["B1"].PutValue("Product");
        worksheet.Cells["C1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("North");
        worksheet.Cells["B2"].PutValue("Widget");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("South");
        worksheet.Cells["B3"].PutValue("Gadget");
        worksheet.Cells["C3"].PutValue(200);

        var pivotTables = worksheet.PivotTables;
        var idx = pivotTables.Add($"={worksheet.Name}!A1:C3", "E1", "TestPivot");
        var pivotTable = pivotTables[idx];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Column, 1);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 2);
        pivotTable.CalculateData();

        return workbook;
    }

    #endregion

    #region Protected Property Tests

    [Fact]
    public void OperationVerb_Returns_Remove()
    {
        var handler = new TestableDeleteFieldHandler();
        Assert.Equal("remove", handler.GetOperationVerb());
    }

    [Fact]
    public void OperationVerbPast_Returns_Removed()
    {
        var handler = new TestableDeleteFieldHandler();
        Assert.Equal("removed", handler.GetOperationVerbPast());
    }

    [Fact]
    public void GetPreposition_Returns_From()
    {
        var handler = new TestableDeleteFieldHandler();
        Assert.Equal("from", handler.GetPrepositionValue());
    }

    private sealed class TestableDeleteFieldHandler : DeleteFieldExcelPivotTableHandler
    {
        public string GetOperationVerb()
        {
            return OperationVerb;
        }

        public string GetOperationVerbPast()
        {
            return OperationVerbPast;
        }

        public string GetPrepositionValue()
        {
            return GetPreposition();
        }
    }

    #endregion

    #region Page Field Tests

    [Fact]
    public void Execute_DeletesPageField()
    {
        var workbook = CreateWorkbookWithPivotTableAndPageField();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "fieldName", "Region" },
            { "fieldType", "page" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("page", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    private static Workbook CreateWorkbookWithPivotTableAndPageField()
    {
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Region");
        worksheet.Cells["B1"].PutValue("Product");
        worksheet.Cells["C1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("North");
        worksheet.Cells["B2"].PutValue("Widget");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("South");
        worksheet.Cells["B3"].PutValue("Gadget");
        worksheet.Cells["C3"].PutValue(200);

        var pivotTables = worksheet.PivotTables;
        var idx = pivotTables.Add($"={worksheet.Name}!A1:C3", "E1", "TestPivot");
        var pivotTable = pivotTables[idx];
        pivotTable.AddFieldToArea(PivotFieldType.Page, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Row, 1);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 2);
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

    #region Generic Exception Path

    [Fact]
    public void Execute_WithInvalidFieldType_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", 0 },
            { "fieldName", "Category" },
            { "fieldType", "invalid" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativePivotTableIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pivotTableIndex", -1 },
            { "fieldName", "Category" },
            { "fieldType", "row" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeSheetIndex_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithPivotTableAndFields();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", -1 },
            { "pivotTableIndex", 0 },
            { "fieldName", "Category" },
            { "fieldType", "row" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("sheet", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
