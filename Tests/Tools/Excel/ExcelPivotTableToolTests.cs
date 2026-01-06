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

    private string CreateWorkbookWithPivotTable(string fileName)
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[0, 0].Value = "Category";
        worksheet.Cells[0, 1].Value = "Sales";
        worksheet.Cells[0, 2].Value = "Region";
        worksheet.Cells[0, 3].Value = "Quantity";
        for (var row = 1; row <= 10; row++)
        {
            worksheet.Cells[row, 0].Value = $"Cat{row % 3}";
            worksheet.Cells[row, 1].Value = row * 100;
            worksheet.Cells[row, 2].Value = $"Region{row % 2}";
            worksheet.Cells[row, 3].Value = row * 10;
        }

        worksheet.PivotTables.Add("A1:D11", "F1", "PivotTable1");
        workbook.Save(workbookPath);
        return workbookPath;
    }

    private string CreateWorkbookWithDataForPivot(string fileName)
    {
        var workbookPath = CreateTestFilePath(fileName);
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets[0];
        worksheet.Cells[0, 0].Value = "Category";
        worksheet.Cells[0, 1].Value = "Sales";
        worksheet.Cells[0, 2].Value = "Region";
        worksheet.Cells[0, 3].Value = "Quantity";
        for (var row = 1; row <= 10; row++)
        {
            worksheet.Cells[row, 0].Value = $"Cat{row % 3}";
            worksheet.Cells[row, 1].Value = row * 100;
            worksheet.Cells[row, 2].Value = $"Region{row % 2}";
            worksheet.Cells[row, 3].Value = row * 10;
        }

        workbook.Save(workbookPath);
        return workbookPath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddPivotTable()
    {
        var workbookPath = CreateWorkbookWithDataForPivot("test_add.xlsx");
        var outputPath = CreateTestFilePath("test_add_output.xlsx");
        var result = _tool.Execute("add", workbookPath, sourceRange: "A1:D11", destCell: "F1", outputPath: outputPath);
        Assert.Contains("added", result);
        using var workbook = new Workbook(outputPath);
        Assert.True(workbook.Worksheets[0].PivotTables.Count > 0);
    }

    [Fact]
    public void Add_WithCustomName_ShouldUseName()
    {
        var workbookPath = CreateWorkbookWithDataForPivot("test_add_name.xlsx");
        var outputPath = CreateTestFilePath("test_add_name_output.xlsx");
        var result = _tool.Execute("add", workbookPath, sourceRange: "A1:D11", destCell: "F1",
            name: "CustomPivot", outputPath: outputPath);
        Assert.Contains("added", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("CustomPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void Get_ShouldReturnPivotTableInfo()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_get.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("count").GetInt32());
        Assert.True(root.TryGetProperty("items", out _));
    }

    [Fact]
    public void Get_NoPivotTables_ShouldReturnEmptyResult()
    {
        var workbookPath = CreateExcelWorkbook("test_get_empty.xlsx");
        var result = _tool.Execute("get", workbookPath);
        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No pivot tables found", json.RootElement.GetProperty("message").GetString());
    }

    [Fact]
    public void Delete_ShouldDeletePivotTable()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_delete.xlsx");
        var outputPath = CreateTestFilePath("test_delete_output.xlsx");
        var result = _tool.Execute("delete", workbookPath, pivotTableIndex: 0, outputPath: outputPath);
        Assert.Contains("deleted", result);
        using var workbook = new Workbook(outputPath);
        Assert.Empty(workbook.Worksheets[0].PivotTables);
    }

    [Fact]
    public void Edit_Name_ShouldEditName()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_name.xlsx");
        var outputPath = CreateTestFilePath("test_edit_name_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            name: "EditedPivot", outputPath: outputPath);
        Assert.Contains("edited", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal("EditedPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void Edit_Style_ShouldApplyStyle()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_style.xlsx");
        var outputPath = CreateTestFilePath("test_edit_style_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            style: "Medium6", outputPath: outputPath);
        Assert.Contains("edited", result);
        using var workbook = new Workbook(outputPath);
        Assert.Equal(PivotTableStyleType.PivotTableStyleMedium6,
            workbook.Worksheets[0].PivotTables[0].PivotTableStyleType);
    }

    [Fact]
    public void Edit_StyleNone_ShouldRemoveStyle()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_style_none.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].PivotTables[0].PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_style_none_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            style: "None", outputPath: outputPath);
        Assert.Contains("edited", result);
        using var workbook = new Workbook(outputPath);
        var pivotTable = workbook.Worksheets[0].PivotTables[0];
        Assert.True(pivotTable.PivotTableStyleType == PivotTableStyleType.None ||
                    pivotTable.PivotTableStyleType == PivotTableStyleType.Custom);
    }

    [Fact]
    public void Edit_ShowRowGrand_ShouldSetRowGrandTotal()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_rowgrand.xlsx");
        var outputPath = CreateTestFilePath("test_edit_rowgrand_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            showRowGrand: false, outputPath: outputPath);
        Assert.Contains("edited", result);
        using var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].PivotTables[0].RowGrand);
    }

    [Fact]
    public void Edit_ShowColumnGrand_ShouldSetColumnGrandTotal()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_colgrand.xlsx");
        var outputPath = CreateTestFilePath("test_edit_colgrand_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            showColumnGrand: false, outputPath: outputPath);
        Assert.Contains("edited", result);
        using var workbook = new Workbook(outputPath);
        Assert.False(workbook.Worksheets[0].PivotTables[0].ColumnGrand);
    }

    [Fact]
    public void Edit_AutoFitColumns_ShouldAutoFit()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_autofit.xlsx");
        var outputPath = CreateTestFilePath("test_edit_autofit_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            autoFitColumns: true, outputPath: outputPath);
        Assert.Contains("edited", result);
    }

    [Fact]
    public void Edit_RefreshData_ShouldRefresh()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_refresh.xlsx");
        var outputPath = CreateTestFilePath("test_edit_refresh_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            refreshData: true, outputPath: outputPath);
        Assert.Contains("edited", result);
    }

    [Fact]
    public void Edit_MultipleOptions_ShouldApplyAll()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_multi.xlsx");
        var outputPath = CreateTestFilePath("test_edit_multi_output.xlsx");
        var result = _tool.Execute("edit", workbookPath, pivotTableIndex: 0,
            name: "StyledPivot", style: "Light10", showRowGrand: true, showColumnGrand: false,
            autoFitColumns: true, refreshData: true, outputPath: outputPath);
        Assert.Contains("edited", result);
        using var workbook = new Workbook(outputPath);
        var pivotTable = workbook.Worksheets[0].PivotTables[0];
        Assert.Equal("StyledPivot", pivotTable.Name);
        Assert.Equal(PivotTableStyleType.PivotTableStyleLight10, pivotTable.PivotTableStyleType);
    }

    [Fact]
    public void AddField_Row_ShouldAddRowField()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_add_field_row.xlsx");
        var outputPath = CreateTestFilePath("test_add_field_row_output.xlsx");
        var result = _tool.Execute("add_field", workbookPath, pivotTableIndex: 0,
            fieldName: "Region", area: "Row", outputPath: outputPath);
        Assert.Contains("added", result);
    }

    [Fact]
    public void AddField_Data_ShouldAddDataField()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_add_field_data.xlsx");
        var outputPath = CreateTestFilePath("test_add_field_data_output.xlsx");
        var result = _tool.Execute("add_field", workbookPath, pivotTableIndex: 0,
            fieldName: "Quantity", fieldType: "Data", function: "Sum", outputPath: outputPath);
        Assert.Contains("added", result);
    }

    [Fact]
    public void DeleteField_ShouldDeleteField()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_delete_field.xlsx");
        var addFieldOutput = CreateTestFilePath("test_delete_field_add.xlsx");
        _tool.Execute("add_field", workbookPath, pivotTableIndex: 0,
            fieldName: "Region", area: "Row", outputPath: addFieldOutput);
        var outputPath = CreateTestFilePath("test_delete_field_output.xlsx");
        var result = _tool.Execute("delete_field", addFieldOutput, pivotTableIndex: 0,
            fieldName: "Region", fieldType: "Row", outputPath: outputPath);
        Assert.Contains("removed", result);
    }

    [Fact]
    public void Refresh_SingleTable_ShouldRefreshData()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_refresh_single.xlsx");
        var outputPath = CreateTestFilePath("test_refresh_single_output.xlsx");
        var result = _tool.Execute("refresh", workbookPath, pivotTableIndex: 0, outputPath: outputPath);
        Assert.Contains("Refreshed", result);
    }

    [Fact]
    public void Refresh_AllTables_ShouldRefreshAll()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_refresh_all.xlsx");
        using (var wb = new Workbook(workbookPath))
        {
            wb.Worksheets[0].PivotTables.Add("A1:D11", "K1", "PivotTable2");
            wb.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_refresh_all_output.xlsx");
        var result = _tool.Execute("refresh", workbookPath, outputPath: outputPath);
        Assert.Contains("Refreshed", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var workbookPath = CreateWorkbookWithDataForPivot($"test_case_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, sourceRange: "A1:D11", destCell: "F1",
            outputPath: outputPath);
        Assert.Contains("added", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var workbookPath = CreateWorkbookWithPivotTable($"test_case_get_{operation}.xlsx");
        var result = _tool.Execute(operation, workbookPath);
        Assert.Contains("count", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var workbookPath = CreateWorkbookWithPivotTable($"test_case_edit_{operation}.xlsx");
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.xlsx");
        var result = _tool.Execute(operation, workbookPath, pivotTableIndex: 0, name: "TestEdit",
            outputPath: outputPath);
        Assert.Contains("edited", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_unknown_op.xlsx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", workbookPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithMissingSourceRange_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_source.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, destCell: "F1"));
        Assert.Contains("sourcerange", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithMissingDestCell_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_add_missing_dest.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sourceRange: "A1:D10"));
        Assert.Contains("destcell", ex.Message.ToLower());
    }

    [Fact]
    public void Add_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithDataForPivot("test_add_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", workbookPath, sheetIndex: 99, sourceRange: "A1:D10", destCell: "F1"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidSheetIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateExcelWorkbook("test_get_invalid_sheet.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", workbookPath, sheetIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidPivotTableIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_delete_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath, pivotTableIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Delete_WithMissingPivotTableIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_delete_missing_index.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", workbookPath));
        Assert.Contains("pivotTableIndex", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidPivotTableIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, pivotTableIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Edit_WithMissingPivotTableIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_missing_index.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, name: "Test"));
        Assert.Contains("pivotTableIndex", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidStyle_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_edit_invalid_style.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", workbookPath, pivotTableIndex: 0, style: "InvalidStyleName"));
        Assert.Contains("Invalid style", ex.Message);
    }

    [Fact]
    public void AddField_WithInvalidPivotTableIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_addfield_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_field", workbookPath, pivotTableIndex: 99, fieldName: "Category", area: "Row"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void AddField_WithMissingFieldName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_addfield_missing_name.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_field", workbookPath, pivotTableIndex: 0, area: "Row"));
        Assert.Contains("fieldName", ex.Message);
    }

    [Fact]
    public void AddField_WithMissingFieldType_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_addfield_missing_type.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_field", workbookPath, pivotTableIndex: 0, fieldName: "Category"));
        Assert.Contains("fieldType", ex.Message);
    }

    [Fact]
    public void AddField_WithFieldNotFound_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_addfield_notfound.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_field", workbookPath, pivotTableIndex: 0, fieldName: "NonExistent", area: "Row"));
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void DeleteField_WithInvalidPivotTableIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_deletefield_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_field", workbookPath, pivotTableIndex: 99, fieldName: "Category", fieldType: "Row"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteField_WithMissingFieldName_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_deletefield_missing_name.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_field", workbookPath, pivotTableIndex: 0, fieldType: "Row"));
        Assert.Contains("fieldName", ex.Message);
    }

    [Fact]
    public void DeleteField_WithMissingFieldType_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_deletefield_missing_type.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete_field", workbookPath, pivotTableIndex: 0, fieldName: "Category"));
        Assert.Contains("fieldType", ex.Message);
    }

    [Fact]
    public void Refresh_WithNoPivotTables_ShouldThrowInvalidOperationException()
    {
        var workbookPath = CreateExcelWorkbook("test_refresh_empty.xlsx");
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tool.Execute("refresh", workbookPath));
        Assert.Contains("No pivot tables found", ex.Message);
    }

    [Fact]
    public void Refresh_WithInvalidPivotTableIndex_ShouldThrowArgumentException()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_refresh_invalid.xlsx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("refresh", workbookPath, pivotTableIndex: 99));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyPath_ShouldThrowException()
    {
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", ""));
    }

    [Fact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get"));
    }

    #endregion

    #region Session

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateWorkbookWithDataForPivot("test_session_add.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add", sessionId: sessionId,
            sourceRange: "A1:D11", destCell: "F1", name: "SessionPivot");
        Assert.Contains("added", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.True(workbook.Worksheets[0].PivotTables.Count > 0);
    }

    [Fact]
    public void Get_WithSessionId_ShouldGetFromMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_get.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Equal(1, JsonDocument.Parse(result).RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_delete.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("delete", sessionId: sessionId, pivotTableIndex: 0);
        Assert.Contains("deleted", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Empty(workbook.Worksheets[0].PivotTables);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_edit.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("edit", sessionId: sessionId, pivotTableIndex: 0, name: "UpdatedPivot");
        Assert.Contains("edited", result);
        var workbook = SessionManager.GetDocument<Workbook>(sessionId);
        Assert.Equal("UpdatedPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public void AddField_WithSessionId_ShouldAddInMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_addfield.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("add_field", sessionId: sessionId, pivotTableIndex: 0,
            fieldName: "Region", area: "Row");
        Assert.Contains("added", result);
    }

    [Fact]
    public void Refresh_WithSessionId_ShouldRefreshInMemory()
    {
        var workbookPath = CreateWorkbookWithPivotTable("test_session_refresh.xlsx");
        var sessionId = OpenSession(workbookPath);
        var result = _tool.Execute("refresh", sessionId: sessionId, pivotTableIndex: 0);
        Assert.Contains("Refreshed", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var workbookPath1 = CreateExcelWorkbook("test_path_file.xlsx");
        var workbookPath2 = CreateWorkbookWithPivotTable("test_session_file.xlsx");
        using (var wb = new Workbook(workbookPath2))
        {
            wb.Worksheets[0].PivotTables[0].Name = "SessionPivotTable";
            wb.Save(workbookPath2);
        }

        var sessionId = OpenSession(workbookPath2);
        var result = _tool.Execute("get", workbookPath1, sessionId);
        Assert.Contains("SessionPivotTable", result);
    }

    #endregion
}