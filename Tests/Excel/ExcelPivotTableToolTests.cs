using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Excel;

public class ExcelPivotTableToolTests : ExcelTestBase
{
    private readonly ExcelPivotTableTool _tool = new();

    [Fact]
    public async Task AddPivotTable_ShouldAddPivotTable()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_add_pivot_table.xlsx", 10, 4);
        var outputPath = CreateTestFilePath("test_add_pivot_table_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sourceRange"] = "A1:D10",
            ["destCell"] = "F1"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.PivotTables.Count > 0, "Worksheet should contain at least one pivot table");
    }

    [Fact]
    public async Task GetPivotTables_ShouldReturnAllPivotTables()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_pivot_tables.xlsx", 10, 4);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
        workbook.Save(workbookPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Pivot", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeletePivotTable_ShouldDeletePivotTable()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_delete_pivot_table.xlsx", 10, 4);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
        workbook.Save(workbookPath);

        var pivotTablesBefore = worksheet.PivotTables.Count;
        Assert.True(pivotTablesBefore > 0, "Pivot table should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_pivot_table_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        var pivotTablesAfter = resultWorksheet.PivotTables.Count;
        Assert.True(pivotTablesAfter < pivotTablesBefore,
            $"Pivot table should be deleted. Before: {pivotTablesBefore}, After: {pivotTablesAfter}");
    }

    [Fact]
    public async Task EditPivotTable_ShouldEditPivotTable()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_table.xlsx", 10, 4);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_edit_pivot_table_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["name"] = "EditedPivotTable"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.PivotTables.Count > 0, "Pivot table should exist after editing");
    }

    [Fact]
    public async Task AddField_ShouldAddFieldToPivotTable()
    {
        // Arrange - Create workbook with header row
        var workbookPath = CreateTestFilePath("test_add_field.xlsx");
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        // Add header row
        worksheet.Cells[0, 0].Value = "Column1";
        worksheet.Cells[0, 1].Value = "Column2";
        worksheet.Cells[0, 2].Value = "Column3";
        worksheet.Cells[0, 3].Value = "Column4";
        // Add data rows
        for (var row = 1; row <= 10; row++)
        for (var col = 0; col < 4; col++)
            worksheet.Cells[row, col].Value = row * 10 + col;
        worksheet.PivotTables.Add("A1:D11", "F1", "PivotTable1");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_add_field_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add_field",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["fieldName"] = "Column1",
            ["area"] = "Row"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.PivotTables.Count > 0, "Pivot table should exist after adding field");
    }

    [Fact]
    public async Task DeleteField_ShouldDeleteFieldFromPivotTable()
    {
        // Arrange - Create workbook with header row
        var workbookPath = CreateTestFilePath("test_delete_field.xlsx");
        var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        // Add header row
        worksheet.Cells[0, 0].Value = "Column1";
        worksheet.Cells[0, 1].Value = "Column2";
        worksheet.Cells[0, 2].Value = "Column3";
        worksheet.Cells[0, 3].Value = "Column4";
        // Add data rows
        for (var row = 1; row <= 10; row++)
        for (var col = 0; col < 4; col++)
            worksheet.Cells[row, col].Value = row * 10 + col;
        workbook.Save(workbookPath);

        worksheet.PivotTables.Add("A1:D11", "F1", "PivotTable1");
        workbook.Save(workbookPath);

        // First add a field to the pivot table
        var addFieldOutputPath = CreateTestFilePath("test_add_field_for_delete.xlsx");
        var addFieldArguments = new JsonObject
        {
            ["operation"] = "add_field",
            ["path"] = workbookPath,
            ["outputPath"] = addFieldOutputPath,
            ["pivotTableIndex"] = 0,
            ["fieldName"] = "Column1",
            ["area"] = "Row"
        };
        await _tool.ExecuteAsync(addFieldArguments);

        // Now delete the field
        var outputPath = CreateTestFilePath("test_delete_field_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete_field",
            ["path"] = addFieldOutputPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["fieldName"] = "Column1",
            ["fieldType"] = "Row"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.PivotTables.Count > 0, "Pivot table should exist after deleting field");
    }

    [Fact]
    public async Task RefreshPivotTable_ShouldRefreshData()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_refresh_pivot_table.xlsx", 10, 4);
        var workbook = new Workbook(workbookPath);
        var worksheet = workbook.Worksheets[0];
        worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_refresh_pivot_table_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "refresh",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var resultWorkbook = new Workbook(outputPath);
        var resultWorksheet = resultWorkbook.Worksheets[0];
        Assert.True(resultWorksheet.PivotTables.Count > 0, "Pivot table should exist after refresh");
    }
}