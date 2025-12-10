using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordEditTableTool : IAsposeTool
{
    public string Description => "Edit formatting of an existing table in a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            tableIndex = new
            {
                type = "number",
                description = "Table index (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            },
            // Table alignment
            alignment = new
            {
                type = "string",
                description = "Table alignment: left, center, right",
                @enum = new[] { "left", "center", "right" }
            },
            // Table width
            width = new
            {
                type = "number",
                description = "Table width in points (or percentage if widthType='percent')"
            },
            widthType = new
            {
                type = "string",
                description = "Width type: auto, points, percent",
                @enum = new[] { "auto", "points", "percent" }
            },
            // Indentation
            leftIndent = new
            {
                type = "number",
                description = "Left indent in points"
            },
            // Cell padding
            cellPaddingTop = new
            {
                type = "number",
                description = "Top cell padding in points"
            },
            cellPaddingBottom = new
            {
                type = "number",
                description = "Bottom cell padding in points"
            },
            cellPaddingLeft = new
            {
                type = "number",
                description = "Left cell padding in points"
            },
            cellPaddingRight = new
            {
                type = "number",
                description = "Right cell padding in points"
            },
            // Cell spacing
            cellSpacing = new
            {
                type = "number",
                description = "Cell spacing in points"
            },
            // Allow auto fit
            allowAutoFit = new
            {
                type = "boolean",
                description = "Allow auto fit columns"
            },
            // Style
            styleName = new
            {
                type = "string",
                description = "Table style name to apply"
            }
        },
        required = new[] { "path", "tableIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
        
        var section = doc.Sections[sectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        
        if (tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range (total tables: {tables.Count})");
        
        var table = tables[tableIndex];
        
        // Apply table alignment
        if (arguments?["alignment"] != null)
        {
            var alignment = arguments["alignment"]?.GetValue<string>() ?? "left";
            table.Alignment = GetTableAlignment(alignment);
        }
        
        // Apply table width (simplified - only support points)
        if (arguments?["width"] != null && arguments?["widthType"]?.GetValue<string>() == "points")
        {
            var width = arguments["width"]?.GetValue<double>();
            if (width.HasValue)
            {
                // Set preferred width in points
                table.PreferredWidth = PreferredWidth.FromPoints(width.Value);
            }
        }
        else if (arguments?["widthType"]?.GetValue<string>() == "auto")
        {
            table.PreferredWidth = PreferredWidth.Auto;
        }
        
        // Apply left indent
        if (arguments?["leftIndent"] != null)
        {
            var leftIndent = arguments["leftIndent"]?.GetValue<double>();
            if (leftIndent.HasValue)
                table.LeftIndent = leftIndent.Value;
        }
        
        // Apply cell padding
        if (arguments?["cellPaddingTop"] != null || arguments?["cellPaddingBottom"] != null ||
            arguments?["cellPaddingLeft"] != null || arguments?["cellPaddingRight"] != null)
        {
            var top = arguments?["cellPaddingTop"]?.GetValue<double>();
            var bottom = arguments?["cellPaddingBottom"]?.GetValue<double>();
            var left = arguments?["cellPaddingLeft"]?.GetValue<double>();
            var right = arguments?["cellPaddingRight"]?.GetValue<double>();
            
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    if (top.HasValue) cell.CellFormat.TopPadding = top.Value;
                    if (bottom.HasValue) cell.CellFormat.BottomPadding = bottom.Value;
                    if (left.HasValue) cell.CellFormat.LeftPadding = left.Value;
                    if (right.HasValue) cell.CellFormat.RightPadding = right.Value;
                }
            }
        }
        
        // Apply cell spacing
        if (arguments?["cellSpacing"] != null)
        {
            var cellSpacing = arguments["cellSpacing"]?.GetValue<double>();
            if (cellSpacing.HasValue)
                table.CellSpacing = cellSpacing.Value;
        }
        
        // Apply allow auto fit
        if (arguments?["allowAutoFit"] != null)
        {
            var allowAutoFit = arguments["allowAutoFit"]?.GetValue<bool>() ?? false;
            table.AllowAutoFit = allowAutoFit;
        }
        
        // Apply table style
        if (arguments?["styleName"] != null)
        {
            var styleName = arguments["styleName"]?.GetValue<string>();
            if (!string.IsNullOrEmpty(styleName))
            {
                try
                {
                    table.Style = doc.Styles[styleName];
                }
                catch
                {
                    // Style not found, ignore
                }
            }
        }
        
        doc.Save(outputPath);
        
        var changes = new List<string>();
        if (arguments?["alignment"] != null) changes.Add($"對齊：{arguments["alignment"]?.GetValue<string>()}");
        if (arguments?["width"] != null) changes.Add($"寬度：{arguments["width"]?.GetValue<double>()}");
        if (arguments?["styleName"] != null) changes.Add($"樣式：{arguments["styleName"]?.GetValue<string>()}");
        
        var changesDesc = changes.Count > 0 ? string.Join("、", changes) : "格式";
        
        return await Task.FromResult($"成功編輯表格 {tableIndex} 的{changesDesc}");
    }
    
    private TableAlignment GetTableAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => TableAlignment.Left,
            "center" => TableAlignment.Center,
            "right" => TableAlignment.Right,
            _ => TableAlignment.Left
        };
    }
}

