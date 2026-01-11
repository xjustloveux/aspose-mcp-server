using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

public class SetBackgroundExcelViewHandler : OperationHandlerBase<Workbook>
{
    public override string Operation => "set_background";

    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var imagePath = parameters.GetOptional<string?>("imagePath");
        var removeBackground = parameters.GetOptional("removeBackground", false);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, sheetIndex);

        if (removeBackground)
        {
            worksheet.BackgroundImage = null;
        }
        else if (!string.IsNullOrEmpty(imagePath))
        {
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");
            var imageBytes = File.ReadAllBytes(imagePath);
            worksheet.BackgroundImage = imageBytes;
        }
        else
        {
            throw new ArgumentException("Either imagePath or removeBackground must be provided");
        }

        MarkModified(context);
        return removeBackground
            ? Success($"Background image removed from sheet {sheetIndex}.")
            : Success($"Background image set for sheet {sheetIndex}.");
    }
}
