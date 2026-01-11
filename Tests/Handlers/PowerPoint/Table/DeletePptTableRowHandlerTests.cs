using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

public class DeletePptTableRowHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptTableRowHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteRow()
    {
        Assert.Equal("delete_row", _handler.Operation);
    }

    #endregion

    #region Slide Index Parameter

    [Fact]
    public void Execute_WithSlideIndex_DeletesFromCorrectSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var tableShapeIndex = AddTableToSlide(pres, 1, 3, 2);
        var table = pres.Slides[1].Shapes[tableShapeIndex] as ITable;
        var initialRowCount = table!.Rows.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", tableShapeIndex },
            { "rowIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialRowCount - 1, table.Rows.Count);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesRow()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DecreasesRowCount()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialRowCount = table!.Rows.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialRowCount - 1, table.Rows.Count);
    }

    [Fact]
    public void Execute_ReturnsDeletedRowIndex()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Row 1", result);
    }

    #endregion

    #region Various Row Indices

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesRowAtVariousIndices(int rowIndex)
    {
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialRowCount = table!.Rows.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", rowIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialRowCount - 1, table.Rows.Count);
    }

    [Fact]
    public void Execute_DeletesFirstRow()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Row 0", result);
    }

    [Fact]
    public void Execute_DeletesLastRow()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rowIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Row 2", result);
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
            { "rowIndex", 0 }
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
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("rowIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 },
            { "rowIndex", 0 }
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
            { "rowIndex", 0 }
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
