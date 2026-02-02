using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

public class EditPptTableCellHandlerTests : PptHandlerTestBase
{
    private readonly EditPptTableCellHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditCell()
    {
        Assert.Equal("edit_cell", _handler.Operation);
    }

    #endregion

    #region Various Cell Positions

    [SkippableTheory]
    [InlineData(0, 0)]
    [InlineData(0, 2)]
    [InlineData(2, 0)]
    [InlineData(2, 2)]
    public void Execute_EditsAnyCell(int row, int col)
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode adds watermark to text");
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", row },
            { "columnIndex", col },
            { "text", $"Cell_{row}_{col}" }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal($"Cell_{row}_{col}", table[col, row].TextFrame.Text);
    }

    #endregion

    #region Slide Index Parameter

    [SkippableFact]
    public void Execute_WithSlideIndex_EditsOnCorrectSlide()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode adds watermark to text");
        var pres = CreatePresentationWithSlides(3);
        var tableShapeIndex = AddTableToSlide(pres, 1, 2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", tableShapeIndex },
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "text", "Edited" }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[1].Shapes[tableShapeIndex] as ITable;
        Assert.NotNull(table);
        Assert.Equal("Edited", table[0, 0].TextFrame.Text);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsCell()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "text", "New Text" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        if (!IsEvaluationMode())
        {
            var table = pres.Slides[0].Shapes[0] as ITable;
            Assert.NotNull(table);
            Assert.Equal("New Text", table[0, 0].TextFrame.Text);
        }
    }

    [Fact]
    public void Execute_ReturnsCellCoordinates()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 1 },
            { "columnIndex", 2 },
            { "text", "Test" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode())
        {
            var table = pres.Slides[0].Shapes[0] as ITable;
            Assert.NotNull(table);
            Assert.Equal("Test", table[2, 1].TextFrame.Text);
        }
    }

    [SkippableFact]
    public void Execute_UpdatesCellContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode adds watermark to text");
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 1 },
            { "columnIndex", 1 },
            { "text", "Updated Content" }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal("Updated Content", table[1, 1].TextFrame.Text);
    }

    [Fact]
    public void Execute_CanSetEmptyText()
    {
        var pres = CreatePresentationWithTableData([["A1", "B1"], ["A2", "B2"]]);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "text", "" }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal("", table[0, 0].TextFrame.Text);
    }

    [SkippableFact]
    public void Execute_WithNonSquareTable_WritesToCorrectRowAndColumn()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode adds watermark to text");
        var pres = CreatePresentationWithTable(3, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 2 },
            { "columnIndex", 0 },
            { "text", "Target" }
        });

        _handler.Execute(context, parameters);

        var table = pres.Slides[0].Shapes[0] as ITable;
        Assert.NotNull(table);
        Assert.Equal("Target", table[0, 2].TextFrame.Text);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "text", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutRowIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 0 },
            { "text", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rowIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 },
            { "text", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 },
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidRowIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 99 },
            { "columnIndex", 0 },
            { "text", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidColumnIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 },
            { "columnIndex", 99 },
            { "text", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNegativeRowIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", -1 },
            { "columnIndex", 0 },
            { "text", "Test" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNonTableShape_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithText("Sample");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 },
            { "columnIndex", 0 },
            { "text", "Test" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not a table", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithTable(int rows, int cols)
    {
        var pres = new Presentation();
        AddTableToSlide(pres, 0, rows, cols);
        return pres;
    }

    private static Presentation CreatePresentationWithTableData(string[][] data)
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var rows = data.Length;
        var cols = data[0].Length;
        var colWidths = Enumerable.Repeat(100.0, cols).ToArray();
        var rowHeights = Enumerable.Repeat(30.0, rows).ToArray();
        var table = slide.Shapes.AddTable(100, 100, colWidths, rowHeights);

        for (var row = 0; row < rows; row++)
        for (var col = 0; col < cols; col++)
            table[col, row].TextFrame.Text = data[row][col];

        return pres;
    }

    private static int AddTableToSlide(Presentation pres, int slideIndex, int rows, int cols)
    {
        var slide = pres.Slides[slideIndex];
        var colWidths = Enumerable.Repeat(100.0, cols).ToArray();
        var rowHeights = Enumerable.Repeat(30.0, rows).ToArray();
        slide.Shapes.AddTable(100, 100, colWidths, rowHeights);
        return slide.Shapes.Count - 1;
    }

    #endregion
}
