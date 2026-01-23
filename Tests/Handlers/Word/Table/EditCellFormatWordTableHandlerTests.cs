using Aspose.Words;
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("entire table", result.Message);
        Assert.Contains("9 cells", result.Message);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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
            { "verticalAlignment", "center" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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
            { "topPadding", 5.0 },
            { "bottomPadding", 5.0 },
            { "leftPadding", 5.0 },
            { "rightPadding", 5.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("row 1", result.Message);
        Assert.Contains("3 cells", result.Message);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("column 1", result.Message);
        Assert.Contains("3 cells", result.Message);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Alignment Options

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("justify")]
    public void Execute_WithAlignment_SetsHorizontalAlignment(string alignment)
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("top")]
    [InlineData("center")]
    [InlineData("bottom")]
    public void Execute_WithVerticalAlignmentFormat_SetsVerticalAlignment(string alignment)
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

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("format", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
