using System.Text.Json.Nodes;
using AsposeMcpServer.Handlers.Excel.DataOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("batch write", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("batch write", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("3", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("sheet 1", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Sheet2Data", workbook.Worksheets[1].Cells["A1"].StringValue);
    }

    [Fact]
    public void Execute_WithNullData_WritesZeroCells()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("0 cells", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
