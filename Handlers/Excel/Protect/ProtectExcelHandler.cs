using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Protect;

/// <summary>
///     Handler for protecting Excel workbook or worksheet with password.
/// </summary>
public class ProtectExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "protect";

    /// <summary>
    ///     Protects workbook or worksheet with password.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: password
    ///     Optional: sheetIndex, protectWorkbook, protectStructure, protectWindows
    /// </param>
    /// <returns>Success message with protection details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var password = parameters.GetRequired<string>("password");
        var sheetIndex = parameters.GetOptional<int?>("sheetIndex");
        var protectWorkbook = parameters.GetOptional("protectWorkbook", false);
        var protectStructure = parameters.GetOptional("protectStructure", true);
        var protectWindows = parameters.GetOptional("protectWindows", false);

        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("password is required for protect operation");

        var workbook = context.Document;

        if (protectWorkbook || (!sheetIndex.HasValue && !protectWorkbook))
        {
            var protectionType = ProtectionType.None;
            if (protectStructure && protectWindows)
                protectionType = ProtectionType.All;
            else if (protectStructure)
                protectionType = ProtectionType.Structure;
            else if (protectWindows)
                protectionType = ProtectionType.Windows;

            if (protectionType != ProtectionType.None)
                workbook.Protect(protectionType, password);
        }
        else if (sheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
            worksheet.Protect(ProtectionType.All, password, null);
        }

        MarkModified(context);

        var target = protectWorkbook ? "workbook" :
            sheetIndex.HasValue ? $"worksheet {sheetIndex.Value}" : "workbook";
        return Success($"Excel {target} protected with password successfully.");
    }
}
