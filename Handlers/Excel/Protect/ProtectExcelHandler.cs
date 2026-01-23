using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Protect;

/// <summary>
///     Handler for protecting Excel workbook or worksheet with password.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    /// <exception cref="ArgumentException">Thrown when password is empty or null.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractProtectExcelParameters(parameters);

        if (string.IsNullOrEmpty(p.Password))
            throw new ArgumentException("password is required for protect operation");

        var workbook = context.Document;

        if (p.ProtectWorkbook || (!p.SheetIndex.HasValue && !p.ProtectWorkbook))
        {
            var protectionType = ProtectionType.None;
            if (p is { ProtectStructure: true, ProtectWindows: true })
                protectionType = ProtectionType.All;
            else if (p.ProtectStructure)
                protectionType = ProtectionType.Structure;
            else if (p.ProtectWindows)
                protectionType = ProtectionType.Windows;

            if (protectionType != ProtectionType.None)
                workbook.Protect(protectionType, p.Password);
        }
        else if (p.SheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex.Value);
            worksheet.Protect(ProtectionType.All, p.Password, null);
        }

        MarkModified(context);

        var target = GetProtectionTarget(p.ProtectWorkbook, p.SheetIndex);
        return new SuccessResult { Message = $"Excel {target} protected with password successfully." };
    }

    /// <summary>
    ///     Gets the protection target description.
    /// </summary>
    /// <param name="protectWorkbook">Whether protecting the workbook.</param>
    /// <param name="sheetIndex">The sheet index if protecting a specific sheet.</param>
    /// <returns>The target description string.</returns>
    private static string GetProtectionTarget(bool protectWorkbook, int? sheetIndex)
    {
        if (protectWorkbook) return "workbook";
        if (sheetIndex.HasValue) return $"worksheet {sheetIndex.Value}";
        return "workbook";
    }

    /// <summary>
    ///     Extracts parameters for ProtectExcel operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static ProtectExcelParameters ExtractProtectExcelParameters(OperationParameters parameters)
    {
        return new ProtectExcelParameters(
            parameters.GetRequired<string>("password"),
            parameters.GetOptional<int?>("sheetIndex"),
            parameters.GetOptional("protectWorkbook", false),
            parameters.GetOptional("protectStructure", true),
            parameters.GetOptional("protectWindows", false)
        );
    }

    /// <summary>
    ///     Parameters for ProtectExcel operation.
    /// </summary>
    /// <param name="Password">The password for protection.</param>
    /// <param name="SheetIndex">The sheet index (optional).</param>
    /// <param name="ProtectWorkbook">Whether to protect the workbook.</param>
    /// <param name="ProtectStructure">Whether to protect the structure.</param>
    /// <param name="ProtectWindows">Whether to protect windows.</param>
    private sealed record ProtectExcelParameters(
        string Password,
        int? SheetIndex,
        bool ProtectWorkbook,
        bool ProtectStructure,
        bool ProtectWindows);
}
