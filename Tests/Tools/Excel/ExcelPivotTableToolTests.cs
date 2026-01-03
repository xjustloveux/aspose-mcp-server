using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Excel;

namespace AsposeMcpServer.Tests.Tools.Excel;

public class ExcelPivotTableToolTests : ExcelTestBase
{
    private readonly ExcelPivotTableTool _tool;

    public ExcelPivotTableToolTests()
    {
        _tool = new ExcelPivotTableTool(SessionManager);
    }

    #region General Tests

    [Fact]
    public void AddPivotTable_ShouldAddPivotTable()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_pivot_table.xlsx", 10, 4);
        var outputPath = CreateTestFilePath("test_add_pivot_table_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            sourceRange: "A1:D10",
            destCell: "F1",
            outputPath: outputPath);
        Assert.Contains("added", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.PivotTables.Count > 0);
    }

    [Fact]
    public void AddPivotTable_WithCustomName_ShouldUseName()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_pivot_custom_name.xlsx", 10, 4);
        var outputPath = CreateTestFilePath("test_add_pivot_custom_name_output.xlsx");
        var result = _tool.Execute(
            "add",
            workbookPath,
            sourceRange: "A1:D10",
            destCell: "F1",
            name: "MyCustomPivot",
            outputPath: outputPath);
        Assert.Contains("MyCustomPivot", result);

        using var workbook = new Workbook(outputPath);
        Assert.Equal("MyCustomPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void AddPivotTable_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_add_pivot_invalid_sheet.xlsx", 10, 4);
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            sheetIndex: 99,
            sourceRange: "A1:D10",
            destCell: "F1"));
    }

    [Fact]
    public void GetPivotTables_ShouldReturnAllPivotTables()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_get_pivot_tables.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(1, root.GetProperty("count").GetInt32());
        Assert.True(root.TryGetProperty("items", out var items));
        Assert.Equal(1, items.GetArrayLength());
    }

    [Fact]
    public void GetPivotTables_NoPivotTables_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_no_pivot_tables.xlsx");
        var result = _tool.Execute(
            "get",
            workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No pivot tables found", root.GetProperty("message").GetString());
    }

    [Fact]
    public void GetPivotTables_InvalidSheetIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_pivot_invalid_sheet.xlsx");
        Assert.Throws<ArgumentException>(() => _tool.Execute(
            "get",
            workbookPath,
            sheetIndex: 99));
    }

    [Fact]
    public void DeletePivotTable_ShouldDeletePivotTable()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_pivot_table.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_pivot_table_output.xlsx");
        var result = _tool.Execute(
            "delete",
            workbookPath,
            pivotTableIndex: 0,
            outputPath: outputPath);
        Assert.Contains("deleted", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].PivotTables);
    }

    [Fact]
    public void DeletePivotTable_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_delete_pivot_invalid_index.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "delete",
            workbookPath,
            pivotTableIndex: 99));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void EditPivotTable_ShouldEditName()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_table.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_table_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            name: "EditedPivotTable",
            outputPath: outputPath);
        Assert.Contains("name=EditedPivotTable", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("EditedPivotTable", resultWorkbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void EditPivotTable_WithRefreshData_ShouldRefresh()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_refresh.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_refresh_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            refreshData: true,
            outputPath: outputPath);
        Assert.Contains("refreshed", result);
    }

    [Fact]
    public void EditPivotTable_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_invalid_index.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 99));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void AddField_ShouldAddFieldToPivotTable()
    {
        // Arrange - Create workbook with header row
        var workbookPath = CreateTestFilePath("test_add_field.xlsx");
        using (var workbook = new Workbook())
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells[0, 0].Value = "Column1";
            worksheet.Cells[0, 1].Value = "Column2";
            worksheet.Cells[0, 2].Value = "Column3";
            worksheet.Cells[0, 3].Value = "Column4";
            for (var row = 1; row <= 10; row++)
            for (var col = 0; col < 4; col++)
                worksheet.Cells[row, col].Value = row * 10 + col;
            worksheet.PivotTables.Add("A1:D11", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_field_output.xlsx");
        var result = _tool.Execute(
            "add_field",
            workbookPath,
            pivotTableIndex: 0,
            fieldName: "Column1",
            area: "Row",
            outputPath: outputPath);
        Assert.Contains("added", result);
        Assert.Contains("Row", result);
    }

    [Fact]
    public void AddField_WithFieldType_ShouldAddField()
    {
        var workbookPath = CreateTestFilePath("test_add_field_type.xlsx");
        using (var workbook = new Workbook())
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells[0, 0].Value = "Category";
            worksheet.Cells[0, 1].Value = "Sales";
            for (var row = 1; row <= 10; row++)
            {
                worksheet.Cells[row, 0].Value = $"Cat{row % 3}";
                worksheet.Cells[row, 1].Value = row * 100;
            }

            worksheet.PivotTables.Add("A1:B11", "D1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_add_field_type_output.xlsx");
        var result = _tool.Execute(
            "add_field",
            workbookPath,
            pivotTableIndex: 0,
            fieldName: "Sales",
            fieldType: "Data",
            function: "Sum",
            outputPath: outputPath);
        Assert.Contains("added", result);
        Assert.Contains("Data", result);
    }

    [Fact]
    public void AddField_InvalidPivotTableIndex_ShouldThrowException()
    {
        var workbookPath = CreateTestFilePath("test_add_field_invalid_pt.xlsx");
        using (var workbook = new Workbook())
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells[0, 0].Value = "Column1";
            worksheet.Cells[0, 1].Value = "Column2";
            for (var row = 1; row <= 5; row++)
            for (var col = 0; col < 2; col++)
                worksheet.Cells[row, col].Value = row * 10 + col;
            worksheet.PivotTables.Add("A1:B6", "D1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add_field",
            workbookPath,
            pivotTableIndex: 99,
            fieldName: "Column1",
            area: "Row"));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void AddField_FieldNotFound_ShouldThrowException()
    {
        var workbookPath = CreateTestFilePath("test_add_field_not_found.xlsx");
        using (var workbook = new Workbook())
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells[0, 0].Value = "Column1";
            worksheet.Cells[0, 1].Value = "Column2";
            for (var row = 1; row <= 5; row++)
            for (var col = 0; col < 2; col++)
                worksheet.Cells[row, col].Value = row * 10 + col;
            worksheet.PivotTables.Add("A1:B6", "D1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add_field",
            workbookPath,
            pivotTableIndex: 0,
            fieldName: "NonExistentColumn",
            area: "Row"));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public void DeleteField_ShouldDeleteFieldFromPivotTable()
    {
        var workbookPath = CreateTestFilePath("test_delete_field.xlsx");
        using (var workbook = new Workbook())
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells[0, 0].Value = "Column1";
            worksheet.Cells[0, 1].Value = "Column2";
            worksheet.Cells[0, 2].Value = "Column3";
            worksheet.Cells[0, 3].Value = "Column4";
            for (var row = 1; row <= 10; row++)
            for (var col = 0; col < 4; col++)
                worksheet.Cells[row, col].Value = row * 10 + col;
            worksheet.PivotTables.Add("A1:D11", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        // First add a field
        var addFieldOutputPath = CreateTestFilePath("test_add_field_for_delete.xlsx");
        _tool.Execute(
            "add_field",
            workbookPath,
            pivotTableIndex: 0,
            fieldName: "Column1",
            area: "Row",
            outputPath: addFieldOutputPath);

        // Now delete the field
        var outputPath = CreateTestFilePath("test_delete_field_output.xlsx");
        var result = _tool.Execute(
            "delete_field",
            addFieldOutputPath,
            pivotTableIndex: 0,
            fieldName: "Column1",
            fieldType: "Row",
            outputPath: outputPath);
        Assert.Contains("removed", result);
    }

    [Fact]
    public void RefreshPivotTable_SingleTable_ShouldRefreshData()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_refresh_pivot_table.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_refresh_pivot_table_output.xlsx");
        var result = _tool.Execute(
            "refresh",
            workbookPath,
            pivotTableIndex: 0,
            outputPath: outputPath);
        Assert.Contains("Refreshed 1 pivot table", result);
    }

    [Fact]
    public void RefreshPivotTable_AllTables_ShouldRefreshAll()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_refresh_all_pivot_tables.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            worksheet.PivotTables.Add("A1:D10", "K1", "PivotTable2");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_refresh_all_pivot_tables_output.xlsx");
        var result = _tool.Execute(
            "refresh",
            workbookPath,
            outputPath: outputPath);
        Assert.Contains("Refreshed 2 pivot table(s)", result);
    }

    [Fact]
    public void RefreshPivotTable_NoPivotTables_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_refresh_no_pivot_tables.xlsx");
        var exception = Assert.Throws<InvalidOperationException>(() => _tool.Execute(
            "refresh",
            workbookPath));
        Assert.Contains("No pivot tables found", exception.Message);
    }

    [Fact]
    public void RefreshPivotTable_InvalidIndex_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_refresh_invalid_index.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "refresh",
            workbookPath,
            pivotTableIndex: 99));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public void EditPivotTable_WithStyle_ShouldApplyStyle()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_style.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_style_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            style: "Medium6",
            outputPath: outputPath);
        Assert.Contains("style=Medium6", result);

        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.Equal(PivotTableStyleType.PivotTableStyleMedium6, pivotTable.PivotTableStyleType);
    }

    [Fact]
    public void EditPivotTable_WithStyleNone_ShouldRemoveStyle()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_style_none.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            var pivotIndex = worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            worksheet.PivotTables[pivotIndex].PivotTableStyleType =
                PivotTableStyleType.PivotTableStyleMedium10;
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_style_none_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            style: "None",
            outputPath: outputPath);
        Assert.Contains("style=None", result);

        // Note: When setting PivotTableStyleType.None, Aspose may report it as Custom
        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.True(pivotTable.PivotTableStyleType == PivotTableStyleType.None ||
                    pivotTable.PivotTableStyleType == PivotTableStyleType.Custom);
    }

    [Fact]
    public void EditPivotTable_WithInvalidStyle_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_invalid_style.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            style: "InvalidStyleName"));
        Assert.Contains("Invalid style", exception.Message);
    }

    [Fact]
    public void EditPivotTable_WithRowGrand_ShouldSetRowGrandTotal()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_row_grand.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_row_grand_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            showRowGrand: false,
            outputPath: outputPath);
        Assert.Contains("showRowGrand=False", result);

        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.False(pivotTable.RowGrand);
    }

    [Fact]
    public void EditPivotTable_WithColumnGrand_ShouldSetColumnGrandTotal()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_col_grand.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_col_grand_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            showColumnGrand: false,
            outputPath: outputPath);
        Assert.Contains("showColumnGrand=False", result);

        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.False(pivotTable.ColumnGrand);
    }

    [Fact]
    public void EditPivotTable_WithAutoFitColumns_ShouldAutoFit()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_autofit.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_autofit_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            autoFitColumns: true,
            outputPath: outputPath);
        Assert.Contains("autoFitColumns", result);
    }

    [Fact]
    public void EditPivotTable_WithMultipleOptions_ShouldApplyAll()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_multi.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_multi_output.xlsx");
        var result = _tool.Execute(
            "edit",
            workbookPath,
            pivotTableIndex: 0,
            name: "StyledPivot",
            style: "Light10",
            showRowGrand: true,
            showColumnGrand: false,
            autoFitColumns: true,
            refreshData: true,
            outputPath: outputPath);
        Assert.Contains("name=StyledPivot", result);
        Assert.Contains("style=Light10", result);
        Assert.Contains("showRowGrand=True", result);
        Assert.Contains("showColumnGrand=False", result);
        Assert.Contains("autoFitColumns", result);
        Assert.Contains("refreshed", result);

        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.Equal("StyledPivot", pivotTable.Name);
        Assert.Equal(PivotTableStyleType.PivotTableStyleLight10, pivotTable.PivotTableStyleType);
        Assert.True(pivotTable.RowGrand);
        Assert.False(pivotTable.ColumnGrand);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbook("test_invalid_op.xlsx");
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "invalid",
            workbookPath));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public void AddPivotTable_MissingSourceRange_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_missing_source.xlsx", 10, 4);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            destCell: "F1"));
        Assert.Contains("sourcerange", exception.Message.ToLower());
    }

    [Fact]
    public void AddPivotTable_MissingDestCell_ShouldThrowException()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_missing_dest.xlsx", 10, 4);
        var exception = Assert.Throws<ArgumentException>(() => _tool.Execute(
            "add",
            workbookPath,
            sourceRange: "A1:D10"));
        Assert.Contains("destcell", exception.Message.ToLower());
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetPivotTables_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_get_pivot.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "SessionPivot");
            workbook.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "get",
            sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
    }

    [Fact]
    public void AddPivotTable_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_add_pivot.xlsx", 10, 4);
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "add",
            sessionId: sessionId,
            sourceRange: "A1:D10",
            destCell: "F1",
            name: "InMemoryPivot");
        Assert.Contains("InMemoryPivot", result);

        // Verify in-memory workbook has the pivot table
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].PivotTables.Count > 0);
        Assert.Equal("InMemoryPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void EditPivotTable_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateExcelWorkbookWithData("test_session_edit_pivot.xlsx", 10, 4);
        using (var wb = new Workbook(workbookPath))
        {
            var worksheet = wb.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "OriginalPivot");
            wb.Save(workbookPath);
        }

        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute(
            "edit",
            sessionId: sessionId,
            pivotTableIndex: 0,
            name: "UpdatedPivot");
        Assert.Contains("name=UpdatedPivot", result);

        // Verify in-memory workbook has updated pivot table
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("UpdatedPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id"));
    }

    #endregion
}