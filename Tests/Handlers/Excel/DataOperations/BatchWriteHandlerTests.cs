using System.Text.Json.Nodes;
using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.DataOperations;

public class BatchWriteHandlerTests : ExcelHandlerTestBase
{
    private readonly BatchWriteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_BatchWrite()
    {
        Assert.Equal("batch_write", _handler.Operation);
    }

    #endregion

    #region Basic Batch Write Operations

    [Fact]
    public void Execute_WithJsonArray_WritesCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var data = new JsonArray
        {
            new JsonObject { ["cell"] = "A1", ["value"] = "Hello" },
            new JsonObject { ["cell"] = "B1", ["value"] = "World" }
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", data }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("batch write", result.ToLower());
        Assert.Contains("2", result);
        Assert.Equal("Hello", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("World", workbook.Worksheets[0].Cells["B1"].StringValue);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithJsonObject_WritesCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var data = new JsonObject
        {
            ["A1"] = "Value1",
            ["B2"] = "Value2",
            ["C3"] = "Value3"
        };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "data", data }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("batch write", result.ToLower());
        Assert.Contains("3", result);
        Assert.Equal("Value1", workbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Value2", workbook.Worksheets[0].Cells["B2"].StringValue);
        Assert.Equal("Value3", workbook.Worksheets[0].Cells["C3"].StringValue);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithSheetIndex_WritesToSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        var context = CreateContext(workbook);
        var data = new JsonObject { ["A1"] = "Sheet2Data" };
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "data", data }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("sheet 1", result.ToLower());
        Assert.Equal("Sheet2Data", workbook.Worksheets[1].Cells["A1"].StringValue);
    }

    [Fact]
    public void Execute_WithNullData_WritesZeroCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0 cells", result.ToLower());
    }

    #endregion
}
