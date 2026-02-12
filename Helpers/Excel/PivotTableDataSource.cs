using Aspose.Cells;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Represents the parsed data source of a pivot table.
/// </summary>
/// <param name="SourceSheet">The worksheet containing the source data.</param>
/// <param name="SourceRange">The range object representing the source data.</param>
public record PivotTableDataSource(Worksheet SourceSheet, Range SourceRange);
