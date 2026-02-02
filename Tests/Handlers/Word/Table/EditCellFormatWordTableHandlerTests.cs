using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Handlers.Word.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Table;

public class EditCellFormatWordTableHandlerTests : WordHandlerTestBase
{
    private readonly EditCellFormatWordTableHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditCellFormat()
    {
        Assert.Equal("edit_cell_format", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithTable(int rows, int cols)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.StartTable();
        for (var i = 0; i < rows; i++)
        {
            for (var j = 0; j < cols; j++)
            {
                builder.InsertCell();
                builder.Write($"R{i}C{j}");
            }

            builder.EndRow();
        }

        builder.EndTable();
        return doc;
    }

    #endregion

    #region Apply To Table

    [Fact]
    public void Execute_WithApplyToTable_FormatsEntireTable()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "applyToTable", true },
            { "backgroundColor", "#FF00FF" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var table = doc.Sections[0].Body.Tables[0];
            foreach (var row in table.Rows.Cast<Row>())
            foreach (var cell in row.Cells.Cast<Cell>())
                Assert.NotEqual(Color.Empty, cell.CellFormat.Shading.BackgroundPatternColor);
        }
    }

    #endregion

    #region Individual Padding

    [Fact]
    public void Execute_WithIndividualPadding_SetsEachPadding()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "paddingTop", 10.0 },
            { "paddingBottom", 8.0 },
            { "paddingLeft", 5.0 },
            { "paddingRight", 5.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            Assert.Equal(10.0, cell.CellFormat.TopPadding);
            Assert.Equal(8.0, cell.CellFormat.BottomPadding);
            Assert.Equal(5.0, cell.CellFormat.LeftPadding);
            Assert.Equal(5.0, cell.CellFormat.RightPadding);
        }
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsCellFormat()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "backgroundColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            Assert.Equal(Color.FromArgb(255, 0, 0).ToArgb(), cell.CellFormat.Shading.BackgroundPatternColor.ToArgb());
        }

        AssertModified(context);
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 1)]
    [InlineData(2, 2)]
    public void Execute_EditsVariousCells(int rowIndex, int colIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", rowIndex },
            { "columnIndex", colIndex },
            { "backgroundColor", "#00FF00" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[rowIndex].Cells[colIndex];
            Assert.Equal(Color.FromArgb(0, 255, 0).ToArgb(), cell.CellFormat.Shading.BackgroundPatternColor.ToArgb());
        }
    }

    #endregion

    #region Format Options

    [Fact]
    public void Execute_WithBackgroundColor_SetsBackgroundColor()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "backgroundColor", "#0000FF" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            Assert.Equal(Color.FromArgb(0, 0, 255).ToArgb(), cell.CellFormat.Shading.BackgroundPatternColor.ToArgb());
        }
    }

    [Fact]
    public void Execute_WithVerticalAlignment_SetsAlignment()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "verticalAlignmentFormat", "center" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            Assert.Equal(CellVerticalAlignment.Center, cell.CellFormat.VerticalAlignment);
        }
    }

    [Fact]
    public void Execute_WithPadding_SetsPadding()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "paddingTop", 5.0 },
            { "paddingBottom", 5.0 },
            { "paddingLeft", 5.0 },
            { "paddingRight", 5.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            Assert.Equal(5.0, cell.CellFormat.TopPadding);
            Assert.Equal(5.0, cell.CellFormat.BottomPadding);
            Assert.Equal(5.0, cell.CellFormat.LeftPadding);
            Assert.Equal(5.0, cell.CellFormat.RightPadding);
        }
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutRowIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 },
            { "backgroundColor", "#FF0000" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rowIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "backgroundColor", "#FF0000" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(99, 0)]
    public void Execute_WithInvalidRowIndex_ThrowsArgumentException(int rowIndex, int colIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", rowIndex },
            { "columnIndex", colIndex },
            { "backgroundColor", "#FF0000" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Theory]
    [InlineData(0, -1)]
    [InlineData(0, 99)]
    public void Execute_WithInvalidColumnIndex_ThrowsArgumentException(int rowIndex, int colIndex)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", rowIndex },
            { "columnIndex", colIndex },
            { "backgroundColor", "#FF0000" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTableIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "backgroundColor", "#FF0000" },
            { "tableIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Table index", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "sectionIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Section index", ex.Message);
    }

    #endregion

    #region Apply To Row

    [Fact]
    public void Execute_WithApplyToRow_FormatsEntireRow()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 1 },
            { "applyToRow", true },
            { "backgroundColor", "#FFFF00" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var row = doc.Sections[0].Body.Tables[0].Rows[1];
            foreach (var cell in row.Cells.Cast<Cell>())
                Assert.NotEqual(Color.Empty, cell.CellFormat.Shading.BackgroundPatternColor);
        }
    }

    [Fact]
    public void Execute_WithApplyToRowWithoutRowIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "applyToRow", true },
            { "backgroundColor", "#FFFF00" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rowIndex is required", ex.Message);
    }

    [Fact]
    public void Execute_WithApplyToRowWithInvalidRowIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 99 },
            { "applyToRow", true },
            { "backgroundColor", "#FFFF00" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Row index", ex.Message);
    }

    #endregion

    #region Apply To Column

    [Fact]
    public void Execute_WithApplyToColumn_FormatsEntireColumn()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 1 },
            { "applyToColumn", true },
            { "backgroundColor", "#00FFFF" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var table = doc.Sections[0].Body.Tables[0];
            foreach (var row in table.Rows.Cast<Row>())
            {
                var cell = row.Cells[1];
                Assert.NotEqual(Color.Empty, cell.CellFormat.Shading.BackgroundPatternColor);
            }
        }
    }

    [Fact]
    public void Execute_WithApplyToColumnWithoutColumnIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "applyToColumn", true },
            { "backgroundColor", "#00FFFF" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex is required", ex.Message);
    }

    #endregion

    #region Text Formatting

    [Fact]
    public void Execute_WithFontName_SetsFontName()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "fontName", "Arial" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            var runs = cell.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            Assert.All(runs, run => Assert.Equal("Arial", run.Font.Name));
        }
    }

    [Fact]
    public void Execute_WithFontSize_SetsFontSize()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "cellFontSize", 14.0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            var runs = cell.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            Assert.All(runs, run => Assert.Equal(14.0, run.Font.Size));
        }
    }

    [Fact]
    public void Execute_WithBoldAndItalic_SetsTextStyle()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "bold", true },
            { "italic", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            foreach (var run in cell.GetChildNodes(NodeType.Run, true).Cast<Run>())
            {
                Assert.True(run.Font.Bold);
                Assert.True(run.Font.Italic);
            }
        }
    }

    [Fact]
    public void Execute_WithColor_SetsTextColor()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "color", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            var runs = cell.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            Assert.All(runs,
                run => Assert.Equal(Color.FromArgb(255, 0, 0).ToArgb(), run.Font.Color.ToArgb()));
        }
    }

    [Fact]
    public void Execute_WithFontNameAsciiAndFarEast_SetsFonts()
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "fontNameAscii", "Arial" },
            { "fontNameFarEast", "MS Gothic" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            foreach (var run in cell.GetChildNodes(NodeType.Run, true).Cast<Run>())
            {
                Assert.Equal("Arial", run.Font.NameAscii);
                Assert.Equal("MS Gothic", run.Font.NameFarEast);
            }
        }
    }

    #endregion

    #region Alignment Options

    [Theory]
    [InlineData("left", ParagraphAlignment.Left)]
    [InlineData("center", ParagraphAlignment.Center)]
    [InlineData("right", ParagraphAlignment.Right)]
    [InlineData("justify", ParagraphAlignment.Justify)]
    public void Execute_WithAlignment_SetsHorizontalAlignment(string alignment, ParagraphAlignment expected)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "alignment", alignment }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            var paragraphs = cell.GetChildNodes(NodeType.Paragraph, true)
                .Cast<Aspose.Words.Paragraph>().ToList();
            Assert.All(paragraphs, para => Assert.Equal(expected, para.ParagraphFormat.Alignment));
        }
    }

    [Theory]
    [InlineData("top", CellVerticalAlignment.Top)]
    [InlineData("center", CellVerticalAlignment.Center)]
    [InlineData("bottom", CellVerticalAlignment.Bottom)]
    public void Execute_WithVerticalAlignmentFormat_SetsVerticalAlignment(string alignment,
        CellVerticalAlignment expected)
    {
        var doc = CreateDocumentWithTable(3, 3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "verticalAlignmentFormat", alignment }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var cell = doc.Sections[0].Body.Tables[0].Rows[0].Cells[0];
            Assert.Equal(expected, cell.CellFormat.VerticalAlignment);
        }
    }

    #endregion
}
