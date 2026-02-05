using Aspose.Cells;
using AsposeMcpServer.Handlers.Excel.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Excel.Table;

public class DeleteExcelTableHandlerTests : ExcelHandlerTestBase
{
    private readonly DeleteExcelTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_ShouldBeDelete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_WithKeepData_ShouldRemoveTableAndKeepData()
    {
        var workbook = CreateWorkbookWithTwoTables();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "keepData", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);
        Assert.Contains("deleted", result.Message);
        Assert.Contains("Data preserved", result.Message);

        // Data should still be present in cells
        Assert.Equal("Name", workbook.Worksheets[0].Cells["A1"].Value);
    }

    [Fact]
    public void Execute_WithoutKeepData_ShouldDeleteTable()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "keepData", false }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(workbook.Worksheets[0].ListObjects);
    }

    [Fact]
    public void Execute_ShouldMarkModified()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "keepData", false }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithMissingTableIndex_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("tableIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ShouldThrow()
    {
        var workbook = CreateWorkbookWithTable();
        var context = CreateContext(workbook);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
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

    private static Workbook CreateWorkbookWithTwoTables()
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        // First table data (A1:C4)
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["B1"].Value = "Value";
        sheet.Cells["C1"].Value = "Count";
        sheet.Cells["A2"].Value = "A";
        sheet.Cells["B2"].Value = 1;
        sheet.Cells["C2"].Value = 10;
        sheet.Cells["A3"].Value = "B";
        sheet.Cells["B3"].Value = 2;
        sheet.Cells["C3"].Value = 20;
        sheet.Cells["A4"].Value = "C";
        sheet.Cells["B4"].Value = 3;
        sheet.Cells["C4"].Value = 30;

        // Second table data (E1:G4)
        sheet.Cells["E1"].Value = "X";
        sheet.Cells["F1"].Value = "Y";
        sheet.Cells["G1"].Value = "Z";
        sheet.Cells["E2"].Value = "D";
        sheet.Cells["F2"].Value = 4;
        sheet.Cells["G2"].Value = 40;
        sheet.Cells["E3"].Value = "E";
        sheet.Cells["F3"].Value = 5;
        sheet.Cells["G3"].Value = 50;
        sheet.Cells["E4"].Value = "F";
        sheet.Cells["F4"].Value = 6;
        sheet.Cells["G4"].Value = 60;

        sheet.ListObjects.Add("A1", "C4", true);
        sheet.ListObjects.Add("E1", "G4", true);
        return workbook;
    }

    #endregion
}
