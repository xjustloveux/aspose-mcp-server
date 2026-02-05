using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeMcpServer.Handlers.Excel.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Table;

public class SetStyleExcelTableHandlerTests : ExcelHandlerTestBase
{
    private readonly SetStyleExcelTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeSetStyle()
    {
        Assert.Equal("set_style", _handler.Operation);
    }

    #endregion

    #region Basic Set Style Operations

    [Fact]
    public void Execute_WithValidStyle_ShouldSetTableStyle()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "styleName", "TableStyleMedium9" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("TableStyleMedium9", result.Message);
        Assert.Equal(TableStyleType.TableStyleMedium9, workbook.Worksheets[0].ListObjects[0].TableStyleType);
    }

    #endregion

    #region ResolveTableStyleType

    [Fact]
    public void ResolveTableStyleType_WithValidName_ShouldReturn()
    {
        var result = SetStyleExcelTableHandler.ResolveTableStyleType("TableStyleMedium9");

        Assert.Equal(TableStyleType.TableStyleMedium9, result);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithTable()
    {
        var workbook = CreateWorkbookWithData(new object[,]
        {
            { "Name", "Value", "Count" },
            { "A", 1, 10 },
            { "B", 2, 20 },
            { "C", 3, 30 }
        });
        workbook.Worksheets[0].ListObjects.Add("A1", "C4", true);
        return workbook;
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingTableIndex_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "styleName", "TableStyleMedium9" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("tableIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithMissingStyleName_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("styleName", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidStyle_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "styleName", "InvalidStyleName" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
