using AsposeMcpServer.Handlers.Excel.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Table;

public class CreateExcelTableHandlerTests : ExcelHandlerTestBase
{
    private readonly CreateExcelTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeCreate()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_WithValidRange_ShouldCreateTable()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Value", "Count" },
            { "A", 1, 10 },
            { "B", 2, 20 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:C3" }
        });

        _handler.Execute(context, parameters);

        Assert.Single(workbook.Worksheets[0].ListObjects);
    }

    [Fact]
    public void Execute_WithName_ShouldSetDisplayName()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Value", "Count" },
            { "A", 1, 10 },
            { "B", 2, 20 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:C3" },
            { "name", "MyTable" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("MyTable", workbook.Worksheets[0].ListObjects[0].DisplayName);
    }

    [Fact]
    public void Execute_WithoutHeaders_ShouldCreateTableWithoutHeaders()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "A", 1, 10 },
            { "B", 2, 20 },
            { "C", 3, 30 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:C3" },
            { "hasHeaders", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("hasHeaders=False", result.Message);
        Assert.Single(workbook.Worksheets[0].ListObjects);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Value", "Count" },
            { "A", 1, 10 },
            { "B", 2, 20 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1:C3" }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidRange_ShouldThrow()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Value" },
            { "A", 1 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithBadRangeFormat_ShouldThrow()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Value" },
            { "A", 1 }
        });
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "range", "A1" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
