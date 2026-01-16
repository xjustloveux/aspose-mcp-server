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
        var editParams = ExtractEditSheetPropertiesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, editParams.SheetIndex);

        if (!string.IsNullOrEmpty(editParams.Name)) worksheet.Name = editParams.Name;

        if (editParams.IsVisible.HasValue) worksheet.IsVisible = editParams.IsVisible.Value;

        if (!string.IsNullOrWhiteSpace(editParams.TabColor))
        {
            var color = ColorHelper.ParseColor(editParams.TabColor);
            worksheet.TabColor = color;
        }

        if (editParams.IsSelected.HasValue && editParams.IsSelected.Value)
            workbook.Worksheets.ActiveSheetIndex = editParams.SheetIndex;

        MarkModified(context);
        return Success($"Sheet {editParams.SheetIndex} properties updated successfully.");
    }

    /// <summary>
    ///     Extracts edit sheet properties parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit sheet properties parameters.</returns>
    private static EditSheetPropertiesParameters ExtractEditSheetPropertiesParameters(OperationParameters parameters)
    {
        return new EditSheetPropertiesParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("name"),
            parameters.GetOptional<bool?>("isVisible"),
            parameters.GetOptional<string?>("tabColor"),
            parameters.GetOptional<bool?>("isSelected")
        );
    }

    /// <summary>
    ///     Record to hold edit sheet properties parameters.
    /// </summary>
    private record EditSheetPropertiesParameters(
        int SheetIndex,
        string? Name,
        bool? IsVisible,
        string? TabColor,
        bool? IsSelected);
}
