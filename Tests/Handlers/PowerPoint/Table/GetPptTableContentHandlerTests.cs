using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Table;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Table;

public class GetPptTableContentHandlerTests : PptHandlerTestBase
{
    private readonly GetPptTableContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetContent()
    {
        Assert.Equal("get_content", _handler.Operation);
    }

    #endregion

    #region Slide Index Parameter

    [Fact]
    public void Execute_WithSlideIndex_GetsFromCorrectSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var tableShapeIndex = AddTableToSlide(pres, 1, 2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", tableShapeIndex }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(1, json.RootElement.GetProperty("slideIndex").GetInt32());
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsTableContent()
    {
        var pres = CreatePresentationWithTable(2, 3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("rowCount", out _));
        Assert.True(json.RootElement.TryGetProperty("columnCount", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectRowCount()
    {
        var pres = CreatePresentationWithTable(4, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(4, json.RootElement.GetProperty("rowCount").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsCorrectColumnCount()
    {
        var pres = CreatePresentationWithTable(2, 5);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(5, json.RootElement.GetProperty("columnCount").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsDataArray()
    {
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("data", out var data));
        Assert.Equal(JsonValueKind.Array, data.ValueKind);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
    }

    [Fact]
    public void Execute_ReturnsShapeIndex()
    {
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("shapeIndex").GetInt32());
    }

    #endregion

    #region Cell Data

    [Fact]
    public void Execute_ReturnsCorrectCellData()
    {
        var pres = CreatePresentationWithTableData([["A1", "B1"], ["A2", "B2"]]);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var data = json.RootElement.GetProperty("data");

        Assert.Equal("A1", data[0][0].GetString());
        Assert.Equal("B1", data[0][1].GetString());
        Assert.Equal("A2", data[1][0].GetString());
        Assert.Equal("B2", data[1][1].GetString());
    }

    [Fact]
    public void Execute_WithEmptyCells_ReturnsEmptyStrings()
    {
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var data = json.RootElement.GetProperty("data");

        Assert.Equal(string.Empty, data[0][0].GetString());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 99 }
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
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("not a table", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithTable(2, 2);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
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
