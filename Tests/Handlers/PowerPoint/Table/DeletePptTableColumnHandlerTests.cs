using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

[SupportedOSPlatform("windows")]
public class DeletePptTableColumnHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptTableColumnHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_DeleteColumn()
    {
        SkipIfNotWindows();
        Assert.Equal("delete_column", _handler.Operation);
    }

    #endregion

    #region Slide Index Parameter

    [SkippableFact]
    public void Execute_WithSlideIndex_DeletesFromCorrectSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var tableShapeIndex = AddTableToSlide(pres, 1, 2, 3);
        var table = pres.Slides[1].Shapes[tableShapeIndex] as ITable;
        var initialColCount = table!.Columns.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", tableShapeIndex },
            { "columnIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialColCount - 1, table.Columns.Count);
    }

    #endregion

    #region Basic Delete Operations

    [SkippableFact]
    public void Execute_DeletesColumn()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialColCount = table!.Columns.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);

        Assert.Equal(initialColCount - 1, table.Columns.Count);
    }

    [SkippableFact]
    public void Execute_DecreasesColumnCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialColCount = table!.Columns.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialColCount - 1, table.Columns.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsDeletedColumnIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.NotNull(table);
        Assert.Equal(2, table.Columns.Count);
    }

    #endregion

    #region Various Column Indices

    [SkippableTheory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesColumnAtVariousIndices(int colIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var initialColCount = table!.Columns.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", colIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialColCount - 1, table.Columns.Count);
    }

    [SkippableFact]
    public void Execute_DeletesFirstColumn()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.NotNull(table);
        Assert.Equal(2, table.Columns.Count);
    }

    [SkippableFact]
    public void Execute_DeletesLastColumn()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var table = pres.Slides[0].Shapes[0] as ITable;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.NotNull(table);
        Assert.Equal(2, table.Columns.Count);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "columnIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithoutColumnIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("columnIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTable(3, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 },
            { "columnIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithNonTableShape_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithText("Sample");
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "columnIndex", 0 }
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
