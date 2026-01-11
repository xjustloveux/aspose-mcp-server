using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Handler for editing worksheet properties in Excel files.
/// </summary>
public class EditSheetPropertiesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit_sheet_properties";

    /// <summary>
    ///     Edits worksheet properties such as name, visibility, tab color, and selection state.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetIndex (0-based)
    ///     Optional: name, isVisible, tabColor, isSelected
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var name = parameters.GetOptional<string?>("name");
        var isVisible = parameters.GetOptional<bool?>("isVisible");
        var tabColor = parameters.GetOptional<string?>("tabColor");
        var isSelected = parameters.GetOptional<bool?>("isSelected");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (!string.IsNullOrEmpty(name)) worksheet.Name = name;

        if (isVisible.HasValue) worksheet.IsVisible = isVisible.Value;

        if (!string.IsNullOrWhiteSpace(tabColor))
        {
            var color = ColorHelper.ParseColor(tabColor);
            worksheet.TabColor = color;
        }

        if (isSelected.HasValue && isSelected.Value) workbook.Worksheets.ActiveSheetIndex = sheetIndex;

        MarkModified(context);
        return Success($"Sheet {sheetIndex} properties updated successfully.");
    }
}
