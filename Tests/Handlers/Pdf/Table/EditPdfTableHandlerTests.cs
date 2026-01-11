using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Table;

public class EditPdfTableHandlerTests : PdfHandlerTestBase
{
    private readonly EditPdfTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreatePdfWithTable(int rows, int columns)
    {
        var document = new Document();
        var page = document.Pages.Add();

        var table = new Aspose.Pdf.Table
        {
            ColumnWidths = string.Join(" ", Enumerable.Repeat("100", columns)),
            DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5F)
        };

        for (var i = 0; i < rows; i++)
        {
            var row = table.Rows.Add();
            for (var j = 0; j < columns; j++)
            {
                var cell = row.Cells.Add();
                cell.Paragraphs.Add(new TextFragment($"Cell {i},{j}"));
            }
        }

        page.Paragraphs.Add(table);
        return document;
    }

    #endregion

    #region Basic Edit Table Operations

    [SkippableFact]
    public void Execute_EditsTableCell()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(3, 3);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "cellRow", 1 },
            { "cellColumn", 1 },
            { "cellValue", "Updated Value" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited", result.ToLower());
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoTables_ThrowsArgumentException()
    {
        var document = CreateEmptyDocument();
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidCellRow_ThrowsArgumentException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "cellRow", 999 },
            { "cellColumn", 0 },
            { "cellValue", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidCellColumn_ThrowsArgumentException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "cellRow", 0 },
            { "cellColumn", 999 },
            { "cellValue", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
