using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Pdf.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

    #region Multiple Tables

    [SkippableFact]
    public void Execute_WithMultipleTables_EditsCorrectTable()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = new Document();
        var page = document.Pages.Add();

        for (var t = 0; t < 3; t++)
        {
            var table = new Aspose.Pdf.Table
            {
                ColumnWidths = "100 100",
                DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5F)
            };
            var row = table.Rows.Add();
            row.Cells.Add().Paragraphs.Add(new TextFragment($"Table {t} Cell 0,0"));
            row.Cells.Add().Paragraphs.Add(new TextFragment($"Table {t} Cell 0,1"));
            page.Paragraphs.Add(table);
        }

        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 1 },
            { "cellRow", 0 },
            { "cellColumn", 0 },
            { "cellValue", "Updated Second Table" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var tables = page.Paragraphs.OfType<Aspose.Pdf.Table>().ToList();
            var editedCell = tables[1].Rows[0].Cells[0];
            var cellText = editedCell.Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(cellText);
            Assert.Equal("Updated Second Table", cellText.Text);

            var untouchedCell = tables[0].Rows[0].Cells[0];
            var untouchedText = untouchedCell.Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(untouchedText);
            Assert.Equal("Table 0 Cell 0,0", untouchedText.Text);
        }

        AssertModified(context);
    }

    #endregion

    #region Default Table Index

    [SkippableFact]
    public void Execute_WithoutTableIndex_EditsFirstTable()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "cellRow", 0 },
            { "cellColumn", 0 },
            { "cellValue", "Updated First Cell" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var table = document.Pages[1].Paragraphs.OfType<Aspose.Pdf.Table>().First();
            var cellText = table.Rows[0].Cells[0].Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(cellText);
            Assert.Equal("Updated First Cell", cellText.Text);
        }

        AssertModified(context);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var table = document.Pages[1].Paragraphs.OfType<Aspose.Pdf.Table>().First();
            var cellText = table.Rows[1].Cells[1].Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(cellText);
            Assert.Equal("Updated Value", cellText.Text);
        }

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

    #region Cell Value Edge Cases

    [SkippableFact]
    public void Execute_WithoutCellValue_DoesNotUpdateCell()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "cellRow", 0 },
            { "cellColumn", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var table = document.Pages[1].Paragraphs.OfType<Aspose.Pdf.Table>().First();
            var cellText = table.Rows[0].Cells[0].Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(cellText);
            Assert.Equal("Cell 0,0", cellText.Text);
        }
    }

    [SkippableFact]
    public void Execute_WithEmptyCellValue_DoesNotUpdateCell()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "cellRow", 0 },
            { "cellColumn", 0 },
            { "cellValue", "" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var table = document.Pages[1].Paragraphs.OfType<Aspose.Pdf.Table>().First();
            var cellText = table.Rows[0].Cells[0].Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(cellText);
            Assert.Equal("Cell 0,0", cellText.Text);
        }
    }

    [SkippableFact]
    public void Execute_WithOnlyCellRow_DoesNotUpdateCell()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "cellRow", 0 },
            { "cellValue", "Test" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var table = document.Pages[1].Paragraphs.OfType<Aspose.Pdf.Table>().First();
            var cellText = table.Rows[0].Cells[0].Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(cellText);
            Assert.Equal("Cell 0,0", cellText.Text);
        }
    }

    [SkippableFact]
    public void Execute_WithOnlyCellColumn_DoesNotUpdateCell()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 0 },
            { "cellColumn", 0 },
            { "cellValue", "Test" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var table = document.Pages[1].Paragraphs.OfType<Aspose.Pdf.Table>().First();
            var cellText = table.Rows[0].Cells[0].Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(cellText);
            Assert.Equal("Cell 0,0", cellText.Text);
        }
    }

    #endregion

    #region Table Index Boundary

    [SkippableFact]
    public void Execute_WithNegativeTableIndex_ThrowsArgumentException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = CreatePdfWithTable(2, 2);
        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", -1 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithLastTableIndex_EditsLastTable()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var document = new Document();
        var page = document.Pages.Add();

        for (var t = 0; t < 2; t++)
        {
            var table = new Aspose.Pdf.Table
            {
                ColumnWidths = "100",
                DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5F)
            };
            var row = table.Rows.Add();
            row.Cells.Add().Paragraphs.Add(new TextFragment($"Table {t}"));
            page.Paragraphs.Add(table);
        }

        var context = CreateContext(document);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "tableIndex", 1 },
            { "cellRow", 0 },
            { "cellColumn", 0 },
            { "cellValue", "Updated Last Table" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode(AsposeLibraryType.Pdf))
        {
            var tables = page.Paragraphs.OfType<Aspose.Pdf.Table>().ToList();
            var lastTableCell = tables[1].Rows[0].Cells[0].Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(lastTableCell);
            Assert.Equal("Updated Last Table", lastTableCell.Text);

            var firstTableCell = tables[0].Rows[0].Cells[0].Paragraphs.OfType<TextFragment>().FirstOrDefault();
            Assert.NotNull(firstTableCell);
            Assert.Equal("Table 0", firstTableCell.Text);
        }

        AssertModified(context);
    }

    #endregion
}
