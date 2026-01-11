using System.Text.Json;
using AsposeMcpServer.Handlers.Excel.Range;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.Range;

public class GetExcelRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly GetExcelRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Include Format

    [Fact]
    public void Execute_WithIncludeFormat_ReturnsFormat()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "includeFormat", true }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("format", out var format));
        Assert.True(format.TryGetProperty("fontName", out _));
        Assert.True(format.TryGetProperty("fontSize", out _));
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyWorkbook()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Original" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Original", workbook.Worksheets[0].Cells["A1"].Value);
        AssertNotModified(context);
    }

    #endregion

    #region Empty Cells

    [Fact]
    public void Execute_WithEmptyCells_ReturnsValue()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("value", out _));
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsRangeData()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A", "B" },
            { "C", "D" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B2" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal("A1:B2", json.RootElement.GetProperty("range").GetString());
        Assert.Equal(4, json.RootElement.GetProperty("count").GetInt32());
        AssertNotModified(context);
    }

    [Theory]
    [InlineData("A1:A1", 1, 1)]
    [InlineData("A1:B2", 2, 2)]
    [InlineData("A1:C3", 3, 3)]
    public void Execute_ReturnsCorrectDimensions(string range, int expectedRows, int expectedCols)
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "1", "2", "3" },
            { "4", "5", "6" },
            { "7", "8", "9" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", range }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(expectedRows, json.RootElement.GetProperty("rowCount").GetInt32());
        Assert.Equal(expectedCols, json.RootElement.GetProperty("columnCount").GetInt32());
    }

    #endregion

    #region Items Array

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Value1", "Value2" }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:B1" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("items", out var items));
        Assert.Equal(JsonValueKind.Array, items.ValueKind);
        Assert.Equal(2, items.GetArrayLength());
    }

    [Fact]
    public void Execute_ItemsContainCellReference()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "Test" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("cell", out var cell));
        Assert.Equal("A1", cell.GetString());
    }

    [Fact]
    public void Execute_ItemsContainValue()
    {
        var workbook = CreateWorkbookWithData(new object[,] { { "TestValue" } });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("value", out var value));
        Assert.Equal("TestValue", value.GetString());
    }

    #endregion

    #region Sheet Index

    [Fact]
    public void Execute_WithSheetIndex_GetsFromSpecificSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Sheet1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = "Sheet2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" },
            { "sheetIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.Contains("Sheet2", items[0].GetProperty("value").GetString());
    }

    [Fact]
    public void Execute_DefaultSheetIndex_GetsFromFirstSheet()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = "FirstSheet";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets[1].Cells["A1"].Value = "SecondSheet";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.Contains("FirstSheet", items[0].GetProperty("value").GetString());
    }

    #endregion

    #region Include Formulas

    [Fact]
    public void Execute_WithIncludeFormulas_ReturnsFormula()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Formula = "=A1*2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1" },
            { "includeFormulas", true }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        Assert.True(items[0].TryGetProperty("formula", out var formula));
        Assert.Contains("A1", formula.GetString());
    }

    [Fact]
    public void Execute_WithoutIncludeFormulas_FormulaIsNull()
    {
        var workbook = CreateEmptyWorkbook();
        workbook.Worksheets[0].Cells["A1"].Value = 10;
        workbook.Worksheets[0].Cells["B1"].Formula = "=A1*2";
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "B1" },
            { "includeFormulas", false }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var items = json.RootElement.GetProperty("items");
        var formula = items[0].GetProperty("formula");
        Assert.Equal(JsonValueKind.Null, formula.ValueKind);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRange_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("InvalidRange")]
    [InlineData("")]
    public void Execute_WithInvalidRange_ThrowsException(string invalidRange)
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", invalidRange }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
