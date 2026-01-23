using AsposeMcpServer.Handlers.Excel.Properties;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Properties;

public class EditSheetPropertiesHandlerTests : ExcelHandlerTestBase
{
    private readonly EditSheetPropertiesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditSheetProperties()
    {
        Assert.Equal("edit_sheet_properties", _handler.Operation);
    }

    #endregion

    #region Basic Edit Sheet Properties Operations

    [Fact]
    public void Execute_RenamesSheet()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "NewName" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("NewName", workbook.Worksheets[0].Name);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsVisibility()
    {
        var workbook = CreateWorkbookWithSheets(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "isVisible", false }
        });

        _handler.Execute(context, parameters);

        Assert.False(workbook.Worksheets[1].IsVisible);
    }

    [Fact]
    public void Execute_SetsTabColor()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tabColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("updated successfully", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_SetsSelectedSheet()
    {
        var workbook = CreateWorkbookWithSheets(3);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 2 },
            { "isSelected", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, workbook.Worksheets.ActiveSheetIndex);
    }

    [Fact]
    public void Execute_WithSheetIndex_UpdatesSpecificSheet()
    {
        var workbook = CreateWorkbookWithSheets(2);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "sheetIndex", 1 },
            { "name", "UpdatedSheet" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("sheet 1", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("UpdatedSheet", workbook.Worksheets[1].Name);
    }

    #endregion
}
