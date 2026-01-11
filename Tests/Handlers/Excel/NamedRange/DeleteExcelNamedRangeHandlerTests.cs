using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.NamedRange;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Excel.NamedRange;

public class DeleteExcelNamedRangeHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelNamedRangeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesNamedRange()
    {
        var workbook = CreateWorkbookWithNamedRange();
        Assert.NotNull(workbook.Worksheets.Names["TestRange"]);
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "TestRange" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Null(workbook.Worksheets.Names["TestRange"]);
        AssertModified(context);
    }

    #endregion

    #region Helper Methods

    private static Workbook CreateWorkbookWithNamedRange()
    {
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells.CreateRange("A1:B5").Name = "TestRange";
        return workbook;
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutName_ThrowsArgumentException()
    {
        var workbook = CreateWorkbookWithNamedRange();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonExistentName_ThrowsArgumentException()
    {
        var workbook = CreateEmptyWorkbook();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "name", "NonExistentRange" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
