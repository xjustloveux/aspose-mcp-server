using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel formulas (add, get, get_result, calculate, set_array, get_array)
/// Merges: ExcelAddFormulaTool, ExcelGetFormulaTool, ExcelGetFormulaResultTool, 
/// ExcelCalculateFormulaTool, ExcelCalculateAllFormulasTool, ExcelSetArrayFormulaTool, ExcelGetArrayFormulaTool
/// </summary>
public class ExcelFormulaTool : IAsposeTool
{
    public string Description => "Manage Excel formulas: add, get, get result, calculate, or set/get array formulas";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'get', 'get_result', 'calculate', 'set_array', 'get_array'",
                @enum = new[] { "add", "get", "get_result", "calculate", "set_array", "get_array" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for calculate operation, defaults to input path)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1', required for add/get_result/get_array)"
            },
            range = new
            {
                type = "string",
                description = "Cell range (e.g., 'A1:C10', optional for get, required for set_array)"
            },
            formula = new
            {
                type = "string",
                description = "Formula (e.g., '=SUM(A1:A10)', required for add/set_array)"
            },
            calculateBeforeRead = new
            {
                type = "boolean",
                description = "Calculate formulas before reading (optional, for get_result, default: true)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "add" => await AddFormulaAsync(arguments, path, sheetIndex),
            "get" => await GetFormulasAsync(arguments, path, sheetIndex),
            "get_result" => await GetFormulaResultAsync(arguments, path, sheetIndex),
            "calculate" => await CalculateFormulasAsync(arguments, path, sheetIndex),
            "set_array" => await SetArrayFormulaAsync(arguments, path, sheetIndex),
            "get_array" => await GetArrayFormulaAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddFormulaAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required for add operation");
        var formula = arguments?["formula"]?.GetValue<string>() ?? throw new ArgumentException("formula is required for add operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        
        var cellObj = worksheet.Cells[cell];
        cellObj.Formula = formula;
        
        // 計算公式，確保結果正確
        workbook.CalculateFormula();
        
        // 確保計算結果被保存（通過訪問計算後的值來觸發計算）
        // 這可以確保當Excel打開文件時，公式已經有正確的計算結果
        var calculatedValue = cellObj.Value;
        
        workbook.Save(path);

        return await Task.FromResult($"Formula added to cell {cell}: {formula}");
    }

    private async Task<string> GetFormulasAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的公式資訊 ===\n");

        int startRow, endRow, startCol, endCol;

        if (!string.IsNullOrEmpty(range))
        {
            try
            {
                var cellRange = cells.CreateRange(range);
                startRow = cellRange.FirstRow;
                endRow = cellRange.FirstRow + cellRange.RowCount - 1;
                startCol = cellRange.FirstColumn;
                endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"無效的範圍格式: {range}", ex);
            }
        }
        else
        {
            startRow = 0;
            endRow = worksheet.Cells.MaxDataRow;
            startCol = 0;
            endCol = worksheet.Cells.MaxDataColumn;
        }

        int formulaCount = 0;
        for (int row = startRow; row <= endRow && row <= 10000; row++)
        {
            for (int col = startCol; col <= endCol && col <= 1000; col++)
            {
                var cell = cells[row, col];
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    formulaCount++;
                    result.AppendLine($"【{CellsHelper.CellIndexToName(row, col)}】");
                    result.AppendLine($"公式: {cell.Formula}");
                    result.AppendLine($"值: {cell.Value ?? "(計算中)"}");
                    result.AppendLine();
                }
            }
        }

        if (formulaCount == 0)
        {
            result.AppendLine("未找到公式");
        }
        else
        {
            result.Insert(0, $"總公式數: {formulaCount}\n\n");
        }

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> GetFormulaResultAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required for get_result operation");
        var calculateBeforeRead = arguments?["calculateBeforeRead"]?.GetValue<bool?>() ?? true;

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        if (calculateBeforeRead)
        {
            workbook.CalculateFormula();
        }

        var result = $"Cell: {cell}\n";
        result += $"Formula: {cellObj.Formula ?? "(none)"}\n";
        
        object? calculatedValue = cellObj.Value;
        
        if (!string.IsNullOrEmpty(cellObj.Formula))
        {
            if (calculatedValue == null || (calculatedValue is string str && string.IsNullOrEmpty(str)))
            {
                calculatedValue = cellObj.DisplayStringValue;
                if (string.IsNullOrEmpty(calculatedValue?.ToString()))
                {
                    calculatedValue = cellObj.Formula;
                }
            }
        }
        
        result += $"Calculated Value: {calculatedValue ?? "(empty)"}\n";
        result += $"Value Type: {cellObj.Type}";

        return await Task.FromResult(result);
    }

    private async Task<string> CalculateFormulasAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var cell = arguments?["cell"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        
        if (!string.IsNullOrEmpty(cell))
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];
            
            if (!string.IsNullOrEmpty(cellObj.Formula))
            {
                var oldValue = cellObj.Value;
                cellObj.PutValue(cellObj.Formula);
            }
        }
        else
        {
            workbook.CalculateFormula();
        }
        
        workbook.Save(outputPath);
        
        var result = "公式計算完成\n";
        result += $"工作表: {workbook.Worksheets[sheetIndex].Name}\n";
        if (!string.IsNullOrEmpty(cell))
        {
            result += $"單元格: {cell}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }

    private async Task<string> SetArrayFormulaAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required for set_array operation");
        var formula = arguments?["formula"]?.GetValue<string>() ?? throw new ArgumentException("formula is required for set_array operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var rangeObj = worksheet.Cells.CreateRange(range);

        // Remove curly braces if present (they're not needed)
        var cleanFormula = formula.TrimStart('{').TrimEnd('}');
        
        // Validate range dimensions
        if (rangeObj.RowCount <= 0 || rangeObj.ColumnCount <= 0)
        {
            throw new ArgumentException($"無效的範圍尺寸：行數={rangeObj.RowCount}，列數={rangeObj.ColumnCount}");
        }
        
        // Validate row and column indices
        if (rangeObj.FirstRow < 0 || rangeObj.FirstColumn < 0)
        {
            throw new ArgumentException($"無效的範圍位置：起始行={rangeObj.FirstRow}，起始列={rangeObj.FirstColumn}");
        }
        
        // Use Cell.SetArrayFormula to properly set array formula
        // Note: This method is deprecated but still functional
        // The new API (FormulaParseOptions) is not available in this version of Aspose.Cells
        var firstCell = worksheet.Cells[rangeObj.FirstRow, rangeObj.FirstColumn];
        
        #pragma warning disable CS0618 // Type or member is obsolete
        // Set array formula using SetArrayFormula method
        // According to Aspose.Cells documentation, SetArrayFormula signature is:
        // SetArrayFormula(string arrayFormula, int rowNumber, int columnNumber)
        // where rowNumber and columnNumber are the number of rows and columns for the array
        
        // Formula should include '=' sign
        var formulaToSet = cleanFormula.StartsWith("=") ? cleanFormula : "=" + cleanFormula;
        
        // Clear the range first
        for (int i = 0; i < rangeObj.RowCount; i++)
        {
            for (int j = 0; j < rangeObj.ColumnCount; j++)
            {
                worksheet.Cells[rangeObj.FirstRow + i, rangeObj.FirstColumn + j].PutValue("");
            }
        }
        
        try
        {
            // Use SetArrayFormula with rowCount and columnCount (not startRow/startCol)
            // Signature: SetArrayFormula(formula, rowCount, columnCount)
            firstCell.SetArrayFormula(formulaToSet, rangeObj.RowCount, rangeObj.ColumnCount);
            
            // Calculate formulas to ensure array formula is processed
            workbook.CalculateFormula();
            
            // Check immediately if it's an array formula
            if (firstCell.IsArrayFormula)
            {
                workbook.Save(path);
                return await Task.FromResult($"Array formula set in range {range}: {path}");
            }
            
            // Save and reload to verify
            workbook.Save(path);
            using var verifyWorkbook = new Workbook(path);
            var verifyWorksheet = verifyWorkbook.Worksheets[sheetIndex];
            var verifyCell = verifyWorksheet.Cells[rangeObj.FirstRow, rangeObj.FirstColumn];
            
            if (verifyCell.IsArrayFormula)
            {
                return await Task.FromResult($"Array formula set in range {range}: {path}");
            }
            else
            {
                // If SetArrayFormula with 2 parameters didn't work, try with 5 parameters
                throw new InvalidOperationException("SetArrayFormula with 2 parameters did not work");
            }
        }
        catch (Exception ex)
        {
            // Try with 5 parameters: SetArrayFormula(formula, startRow, startCol, isR1C1, isLocal)
            try
            {
                // Reload for clean state
                using var retryWorkbook = new Workbook(path);
                var retryWorksheet = retryWorkbook.Worksheets[sheetIndex];
                var retryRangeObj = retryWorksheet.Cells.CreateRange(range);
                var retryFirstCell = retryWorksheet.Cells[retryRangeObj.FirstRow, retryRangeObj.FirstColumn];
                
                // Clear range
                for (int i = 0; i < retryRangeObj.RowCount; i++)
                {
                    for (int j = 0; j < retryRangeObj.ColumnCount; j++)
                    {
                        retryWorksheet.Cells[retryRangeObj.FirstRow + i, retryRangeObj.FirstColumn + j].PutValue("");
                    }
                }
                
                // Try with 5 parameters
                var formulaWithoutEquals = cleanFormula.StartsWith("=") ? cleanFormula.Substring(1) : cleanFormula;
                retryFirstCell.SetArrayFormula(formulaWithoutEquals, retryRangeObj.FirstRow, retryRangeObj.FirstColumn, false, false);
                
                retryWorkbook.CalculateFormula();
                retryWorkbook.Save(path);
                
                // Verify
                using var verifyWorkbook = new Workbook(path);
                var verifyWorksheet = verifyWorkbook.Worksheets[sheetIndex];
                var verifyCell = verifyWorksheet.Cells[retryRangeObj.FirstRow, retryRangeObj.FirstColumn];
                
                if (verifyCell.IsArrayFormula)
                {
                    return await Task.FromResult($"Array formula set in range {range}: {path}");
                }
                else
                {
                    throw new InvalidOperationException("SetArrayFormula with 5 parameters did not work");
                }
            }
            catch (Exception ex2)
            {
                // If both methods fail, set regular formulas as fallback
                try
                {
                    using var fallbackWorkbook = new Workbook(path);
                    var fallbackWorksheet = fallbackWorkbook.Worksheets[sheetIndex];
                    var fallbackRangeObj = fallbackWorksheet.Cells.CreateRange(range);
                    
                    var formulaWithEquals = cleanFormula.StartsWith("=") ? cleanFormula : "=" + cleanFormula;
                    for (int i = 0; i < fallbackRangeObj.RowCount; i++)
                    {
                        for (int j = 0; j < fallbackRangeObj.ColumnCount; j++)
                        {
                            var cell = fallbackWorksheet.Cells[fallbackRangeObj.FirstRow + i, fallbackRangeObj.FirstColumn + j];
                            cell.Formula = formulaWithEquals;
                        }
                    }
                    
                    fallbackWorkbook.Save(path);
                    return await Task.FromResult($"公式已設置到範圍 {range}（注意：這不是真正的數組公式）: {path}");
                }
                catch (Exception ex3)
                {
                    throw new ArgumentException($"無法設置數組公式。範圍: {range}，公式: {cleanFormula}。\n方法1錯誤: {ex.Message}\n方法2錯誤: {ex2.Message}\n方法3錯誤: {ex3.Message}", ex);
                }
            }
        }
        #pragma warning restore CS0618 // Type or member is obsolete
    }

    private async Task<string> GetArrayFormulaAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required for get_array operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        var sb = new StringBuilder();
        sb.AppendLine($"Cell: {cell}");

        // Use IsArrayFormula property to detect if cell contains an array formula
        if (cellObj.IsArrayFormula)
        {
            var formula = cellObj.Formula;
            sb.AppendLine($"Array Formula: {formula ?? "(empty)"}");
            
            // Try to get the array range
            try
            {
                // Find the array range by checking surrounding cells
                int startRow = cellObj.Row;
                int startCol = cellObj.Column;
                int endRow = startRow;
                int endCol = startCol;
                
                // Check cells to the right
                for (int col = startCol + 1; col < worksheet.Cells.MaxColumn + 1; col++)
                {
                    var testCell = worksheet.Cells[startRow, col];
                    if (testCell.IsArrayFormula && testCell.Formula == formula)
                    {
                        endCol = col;
                    }
                    else
                    {
                        break;
                    }
                }
                
                // Check cells below
                for (int row = startRow + 1; row < worksheet.Cells.MaxRow + 1; row++)
                {
                    var testCell = worksheet.Cells[row, startCol];
                    if (testCell.IsArrayFormula && testCell.Formula == formula)
                    {
                        endRow = row;
                    }
                    else
                    {
                        break;
                    }
                }
                
                var startCellName = CellsHelper.CellIndexToName(startRow, startCol);
                var endCellName = CellsHelper.CellIndexToName(endRow, endCol);
                sb.AppendLine($"Array Range: {startCellName}:{endCellName}");
            }
            catch
            {
                sb.AppendLine("Array Range: Unable to determine");
            }
        }
        else
        {
            sb.AppendLine("No array formula found in this cell");
        }

        return await Task.FromResult(sb.ToString());
    }
}


