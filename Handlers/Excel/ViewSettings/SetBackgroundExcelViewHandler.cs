using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting or removing background image in Excel worksheets.
/// </summary>
public class SetBackgroundExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_background";

    /// <summary>
    ///     Sets or removes the background image for a worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), imagePath, removeBackground (default: false)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when neither imagePath nor removeBackground is provided.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the image file is not found.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetBackgroundParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);

        if (p.RemoveBackground)
        {
            worksheet.BackgroundImage = null;
        }
        else if (!string.IsNullOrEmpty(p.ImagePath))
        {
            if (!File.Exists(p.ImagePath))
                throw new FileNotFoundException($"Image file not found: {p.ImagePath}");
            var imageBytes = File.ReadAllBytes(p.ImagePath);
            worksheet.BackgroundImage = imageBytes;
        }
        else
        {
            throw new ArgumentException("Either imagePath or removeBackground must be provided");
        }

        MarkModified(context);
        return p.RemoveBackground
            ? Success($"Background image removed from sheet {p.SheetIndex}.")
            : Success($"Background image set for sheet {p.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts set background parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetBackgroundParameters record containing all extracted values.</returns>
    private static SetBackgroundParameters ExtractSetBackgroundParameters(OperationParameters parameters)
    {
        return new SetBackgroundParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("imagePath"),
            parameters.GetOptional("removeBackground", false)
        );
    }

    /// <summary>
    ///     Record containing parameters for set background operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="ImagePath">The path to the background image file.</param>
    /// <param name="RemoveBackground">Whether to remove the background image.</param>
    private record SetBackgroundParameters(
        int SheetIndex,
        string? ImagePath,
        bool RemoveBackground);
}
