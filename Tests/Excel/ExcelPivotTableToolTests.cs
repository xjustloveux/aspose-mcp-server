using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Pivot;
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("added", result);

        using var workbook = new Workbook(outputPath);
        var worksheet = workbook.Worksheets[0];
        Assert.True(worksheet.PivotTables.Count > 0);
    }

    [Fact]
    public async Task AddPivotTable_WithCustomName_ShouldUseName()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_add_pivot_custom_name.xlsx", 10, 4);
        var outputPath = CreateTestFilePath("test_add_pivot_custom_name_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["sourceRange"] = "A1:D10",
            ["destCell"] = "F1",
            ["name"] = "MyCustomPivot"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("MyCustomPivot", result);

        using var workbook = new Workbook(outputPath);
        Assert.Equal("MyCustomPivot", workbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public async Task AddPivotTable_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_add_pivot_invalid_sheet.xlsx", 10, 4);
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = workbookPath,
            ["sourceRange"] = "A1:D10",
            ["destCell"] = "F1",
            ["sheetIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetPivotTables_ShouldReturnAllPivotTables()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_get_pivot_tables.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(1, root.GetProperty("count").GetInt32());
        Assert.True(root.TryGetProperty("items", out var items));
        Assert.Equal(1, items.GetArrayLength());
    }

    [Fact]
    public async Task GetPivotTables_NoPivotTables_ShouldReturnEmptyResult()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_no_pivot_tables.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonDocument.Parse(result);
        var root = json.RootElement;

        Assert.Equal(0, root.GetProperty("count").GetInt32());
        Assert.Equal("No pivot tables found", root.GetProperty("message").GetString());
    }

    [Fact]
    public async Task GetPivotTables_InvalidSheetIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_get_pivot_invalid_sheet.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = workbookPath,
            ["sheetIndex"] = 99
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task DeletePivotTable_ShouldDeletePivotTable()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_delete_pivot_table.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_delete_pivot_table_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("deleted", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Empty(resultWorkbook.Worksheets[0].PivotTables);
    }

    [Fact]
    public async Task DeletePivotTable_InvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_delete_pivot_invalid_index.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = workbookPath,
            ["pivotTableIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task EditPivotTable_ShouldEditName()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_table.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("name=EditedPivotTable", result);

        using var resultWorkbook = new Workbook(outputPath);
        Assert.Equal("EditedPivotTable", resultWorkbook.Worksheets[0].PivotTables[0].Name);
    }

    [Fact]
    public async Task EditPivotTable_WithRefreshData_ShouldRefresh()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_refresh.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_refresh_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["refreshData"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("refreshed", result);
    }

    [Fact]
    public async Task EditPivotTable_InvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_invalid_index.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["pivotTableIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task AddField_ShouldAddFieldToPivotTable()
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("added", result);
        Assert.Contains("Row", result);
    }

    [Fact]
    public async Task AddField_WithFieldType_ShouldAddField()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "add_field",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["fieldName"] = "Sales",
            ["fieldType"] = "Data",
            ["function"] = "Sum"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("added", result);
        Assert.Contains("Data", result);
    }

    [Fact]
    public async Task AddField_InvalidPivotTableIndex_ShouldThrowException()
    {
        // Arrange
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

        var arguments = new JsonObject
        {
            ["operation"] = "add_field",
            ["path"] = workbookPath,
            ["pivotTableIndex"] = 99,
            ["fieldName"] = "Column1",
            ["area"] = "Row"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task AddField_FieldNotFound_ShouldThrowException()
    {
        // Arrange
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

        var arguments = new JsonObject
        {
            ["operation"] = "add_field",
            ["path"] = workbookPath,
            ["pivotTableIndex"] = 0,
            ["fieldName"] = "NonExistentColumn",
            ["area"] = "Row"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("not found", exception.Message);
    }

    [Fact]
    public async Task DeleteField_ShouldDeleteFieldFromPivotTable()
    {
        // Arrange
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
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("removed", result);
    }

    [Fact]
    public async Task RefreshPivotTable_SingleTable_ShouldRefreshData()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_refresh_pivot_table.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_refresh_pivot_table_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "refresh",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Refreshed 1 pivot table", result);
    }

    [Fact]
    public async Task RefreshPivotTable_AllTables_ShouldRefreshAll()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_refresh_all_pivot_tables.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            worksheet.PivotTables.Add("A1:D10", "K1", "PivotTable2");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_refresh_all_pivot_tables_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "refresh",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Refreshed 2 pivot table(s)", result);
    }

    [Fact]
    public async Task RefreshPivotTable_NoPivotTables_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_refresh_no_pivot_tables.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "refresh",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("No pivot tables found", exception.Message);
    }

    [Fact]
    public async Task RefreshPivotTable_InvalidIndex_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_refresh_invalid_index.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "refresh",
            ["path"] = workbookPath,
            ["pivotTableIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("out of range", exception.Message);
    }

    [Fact]
    public async Task ExecuteAsync_InvalidOperation_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbook("test_invalid_op.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "invalid",
            ["path"] = workbookPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task EditPivotTable_WithStyle_ShouldApplyStyle()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_style.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_style_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["style"] = "Medium6"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("style=Medium6", result);

        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.Equal(PivotTableStyleType.PivotTableStyleMedium6, pivotTable.PivotTableStyleType);
    }

    [Fact]
    public async Task EditPivotTable_WithStyleNone_ShouldRemoveStyle()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["style"] = "None"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("style=None", result);

        // Note: When setting PivotTableStyleType.None, Aspose may report it as Custom
        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.True(pivotTable.PivotTableStyleType == PivotTableStyleType.None ||
                    pivotTable.PivotTableStyleType == PivotTableStyleType.Custom);
    }

    [Fact]
    public async Task EditPivotTable_WithInvalidStyle_ShouldThrowException()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_invalid_style.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["pivotTableIndex"] = 0,
            ["style"] = "InvalidStyleName"
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Invalid style", exception.Message);
    }

    [Fact]
    public async Task EditPivotTable_WithRowGrand_ShouldSetRowGrandTotal()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_row_grand.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_row_grand_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["showRowGrand"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("showRowGrand=False", result);

        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.False(pivotTable.RowGrand);
    }

    [Fact]
    public async Task EditPivotTable_WithColumnGrand_ShouldSetColumnGrandTotal()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_col_grand.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_col_grand_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["showColumnGrand"] = false
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("showColumnGrand=False", result);

        using var resultWorkbook = new Workbook(outputPath);
        var pivotTable = resultWorkbook.Worksheets[0].PivotTables[0];
        Assert.False(pivotTable.ColumnGrand);
    }

    [Fact]
    public async Task EditPivotTable_WithAutoFitColumns_ShouldAutoFit()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_autofit.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_autofit_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["autoFitColumns"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("autoFitColumns", result);
    }

    [Fact]
    public async Task EditPivotTable_WithMultipleOptions_ShouldApplyAll()
    {
        // Arrange
        var workbookPath = CreateExcelWorkbookWithData("test_edit_pivot_multi.xlsx", 10, 4);
        using (var workbook = new Workbook(workbookPath))
        {
            var worksheet = workbook.Worksheets[0];
            worksheet.PivotTables.Add("A1:D10", "F1", "PivotTable1");
            workbook.Save(workbookPath);
        }

        var outputPath = CreateTestFilePath("test_edit_pivot_multi_output.xlsx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = workbookPath,
            ["outputPath"] = outputPath,
            ["pivotTableIndex"] = 0,
            ["name"] = "StyledPivot",
            ["style"] = "Light10",
            ["showRowGrand"] = true,
            ["showColumnGrand"] = false,
            ["autoFitColumns"] = true,
            ["refreshData"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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
}