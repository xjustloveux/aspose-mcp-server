using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.List;

namespace AsposeMcpServer.Tests.Helpers.Word;

public class WordListHelperTests
{
    #region ParseItems Tests

    [Fact]
    public void ParseItems_WithEmptyArray_ThrowsArgumentException()
    {
        var items = new JsonArray();

        var ex = Assert.Throws<ArgumentException>(() => WordListHelper.ParseItems(items));
        Assert.Contains("empty", ex.Message);
    }

    [Fact]
    public void ParseItems_WithStringValues_ReturnsLevel0()
    {
        var items = new JsonArray("Item 1", "Item 2", "Item 3");

        var result = WordListHelper.ParseItems(items);

        Assert.Equal(3, result.Count);
        Assert.All(result, item => Assert.Equal(0, item.Level));
        Assert.Equal("Item 1", result[0].Text);
        Assert.Equal("Item 2", result[1].Text);
        Assert.Equal("Item 3", result[2].Text);
    }

    [Fact]
    public void ParseItems_WithNullItems_SkipsNulls()
    {
        var items = new JsonArray("Item 1", null, "Item 3");

        var result = WordListHelper.ParseItems(items);

        Assert.Equal(2, result.Count);
        Assert.Equal("Item 1", result[0].Text);
        Assert.Equal("Item 3", result[1].Text);
    }

    [Fact]
    public void ParseItems_WithAllNullItems_ThrowsArgumentException()
    {
        var items = new JsonArray(null, null, null);

        var ex = Assert.Throws<ArgumentException>(() => WordListHelper.ParseItems(items));
        Assert.Contains("No valid list items", ex.Message);
    }

    [Fact]
    public void ParseItems_WithJsonObjects_ReturnsTextAndLevel()
    {
        var items = new JsonArray(
            new JsonObject { ["text"] = "Level 0 item", ["level"] = 0 },
            new JsonObject { ["text"] = "Level 1 item", ["level"] = 1 },
            new JsonObject { ["text"] = "Level 2 item", ["level"] = 2 }
        );

        var result = WordListHelper.ParseItems(items);

        Assert.Equal(3, result.Count);
        Assert.Equal("Level 0 item", result[0].Text);
        Assert.Equal(0, result[0].Level);
        Assert.Equal("Level 1 item", result[1].Text);
        Assert.Equal(1, result[1].Level);
        Assert.Equal("Level 2 item", result[2].Text);
        Assert.Equal(2, result[2].Level);
    }

    [Fact]
    public void ParseItems_WithMissingText_ThrowsArgumentException()
    {
        var items = new JsonArray(
            new JsonObject { ["level"] = 0 }
        );

        var ex = Assert.Throws<ArgumentException>(() => WordListHelper.ParseItems(items));
        Assert.Contains("text", ex.Message);
    }

    [Fact]
    public void ParseItems_WithLevelAboveRange_ClampsTo8()
    {
        var items = new JsonArray(
            new JsonObject { ["text"] = "High level", ["level"] = 99 }
        );

        var result = WordListHelper.ParseItems(items);

        Assert.Single(result);
        Assert.Equal(8, result[0].Level);
    }

    [Fact]
    public void ParseItems_WithLevelBelowRange_ClampsTo0()
    {
        var items = new JsonArray(
            new JsonObject { ["text"] = "Negative level", ["level"] = -5 }
        );

        var result = WordListHelper.ParseItems(items);

        Assert.Single(result);
        Assert.Equal(0, result[0].Level);
    }

    [Fact]
    public void ParseItems_WithJsonObjectWithoutLevel_DefaultsTo0()
    {
        var items = new JsonArray(
            new JsonObject { ["text"] = "No level" }
        );

        var result = WordListHelper.ParseItems(items);

        Assert.Single(result);
        Assert.Equal("No level", result[0].Text);
        Assert.Equal(0, result[0].Level);
    }

    [Fact]
    public void ParseItems_WithMixedFormats_ParsesCorrectly()
    {
        var items = new JsonArray(
            "Simple string",
            new JsonObject { ["text"] = "Object item", ["level"] = 2 }
        );

        var result = WordListHelper.ParseItems(items);

        Assert.Equal(2, result.Count);
        Assert.Equal("Simple string", result[0].Text);
        Assert.Equal(0, result[0].Level);
        Assert.Equal("Object item", result[1].Text);
        Assert.Equal(2, result[1].Level);
    }

    #endregion

    #region BuildListFormatInfo Tests

    [Fact]
    public void BuildListFormatInfo_WithNonListParagraph_ReturnsNonListInfo()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Not a list item");

        var para = doc.FirstSection.Body.FirstParagraph;
        var indices = new Dictionary<(int listId, int paraIndex), int>();

        var result = WordListHelper.BuildListFormatInfo(para, 0, indices);

        var anonObj = result;
        var isListItem = anonObj.GetType().GetProperty("isListItem")?.GetValue(anonObj);
        Assert.NotNull(isListItem);
        Assert.False((bool)isListItem);
    }

    [Fact]
    public void BuildListFormatInfo_WithListParagraph_ReturnsListInfo()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;
        builder.Writeln("List item 1");

        var para = doc.FirstSection.Body.FirstParagraph;
        var indices = new Dictionary<(int listId, int paraIndex), int>
        {
            { (list.ListId, 0), 0 }
        };

        var result = WordListHelper.BuildListFormatInfo(para, 0, indices);

        Assert.IsType<Dictionary<string, object?>>(result);
        var dict = (Dictionary<string, object?>)result;
        Assert.True((bool)dict["isListItem"]!);
        Assert.Equal(0, dict["listLevel"]);
    }

    #endregion

    #region BuildListFormatSingleResult Tests

    [Fact]
    public void BuildListFormatSingleResult_WithNonListParagraph_ReturnsNonListResult()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Not a list item");

        var para = doc.FirstSection.Body.FirstParagraph;
        var indices = new Dictionary<(int listId, int paraIndex), int>();

        var result = WordListHelper.BuildListFormatSingleResult(para, 0, indices);

        Assert.IsType<GetWordListFormatSingleResult>(result);
        Assert.False(result.IsListItem);
        Assert.NotNull(result.Note);
        Assert.Contains("not a list item", result.Note, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BuildListFormatSingleResult_WithListParagraph_ReturnsListResult()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var list = doc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;
        builder.Writeln("Numbered item");

        var para = doc.FirstSection.Body.FirstParagraph;
        var indices = new Dictionary<(int listId, int paraIndex), int>
        {
            { (list.ListId, 0), 0 }
        };

        var result = WordListHelper.BuildListFormatSingleResult(para, 0, indices);

        Assert.True(result.IsListItem);
        Assert.Equal(0, result.ListLevel);
        Assert.Equal(list.ListId, result.ListId);
        Assert.Equal(0, result.ListItemIndex);
        Assert.NotNull(result.ListLevelFormat);
    }

    #endregion

    #region BuildListParagraphInfo Tests

    [Fact]
    public void BuildListParagraphInfo_WithNonListParagraph_ReturnsNonListInfo()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Plain paragraph");

        var para = doc.FirstSection.Body.FirstParagraph;
        var indices = new Dictionary<(int listId, int paraIndex), int>();

        var result = WordListHelper.BuildListParagraphInfo(para, 0, indices);

        Assert.IsType<ListParagraphInfo>(result);
        Assert.False(result.IsListItem);
        Assert.Contains("Plain paragraph", result.ContentPreview);
    }

    [Fact]
    public void BuildListParagraphInfo_WithListParagraph_ReturnsListInfo()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var list = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;
        builder.Writeln("Bullet item");

        var para = doc.FirstSection.Body.FirstParagraph;
        var indices = new Dictionary<(int listId, int paraIndex), int>
        {
            { (list.ListId, 0), 0 }
        };

        var result = WordListHelper.BuildListParagraphInfo(para, 0, indices);

        Assert.True(result.IsListItem);
        Assert.Equal(list.ListId, result.ListId);
        Assert.Equal(0, result.ListItemIndex);
        Assert.NotNull(result.ListLevelFormat);
    }

    [Fact]
    public void BuildListFormatInfo_WithLongText_TruncatesPreview()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        var longText = new string('A', 100);
        builder.Write(longText);

        var para = doc.FirstSection.Body.FirstParagraph;
        var indices = new Dictionary<(int listId, int paraIndex), int>();

        var result = WordListHelper.BuildListFormatInfo(para, 0, indices);

        var previewProp = result.GetType().GetProperty("contentPreview")?.GetValue(result) as string;
        Assert.NotNull(previewProp);
        Assert.True(previewProp.Length <= 53);
        Assert.EndsWith("...", previewProp);
    }

    #endregion
}
