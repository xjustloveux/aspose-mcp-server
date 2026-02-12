using System.Reflection;
using AsposeMcpServer.Tools.Conversion;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tests.Tools;

public class ToolAnnotationsTests
{
    private static IEnumerable<Type> GetAllToolTypes()
    {
        return typeof(ConvertDocumentTool).Assembly.GetTypes()
            .Where(t => t.GetCustomAttribute<McpServerToolTypeAttribute>() != null);
    }

    private static IEnumerable<(Type Type, MethodInfo Method, McpServerToolAttribute Attr)> GetAllTools()
    {
        foreach (var type in GetAllToolTypes())
        {
            var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() != null);

            foreach (var method in methods)
            {
                var attr = method.GetCustomAttribute<McpServerToolAttribute>()!;
                yield return (type, method, attr);
            }
        }
    }

    [Fact]
    public void AllTools_ShouldHaveOpenWorldFalse()
    {
        var tools = GetAllTools().ToList();

        Assert.NotEmpty(tools);

        foreach (var (type, method, attr) in tools)
            Assert.False(
                attr.OpenWorld,
                $"Tool '{attr.Name}' in {type.Name}.{method.Name} should have OpenWorld = false (local file operation)");
    }

    [Fact]
    public void AllTools_ShouldHaveTitle()
    {
        var tools = GetAllTools().ToList();

        Assert.NotEmpty(tools);

        foreach (var (type, method, attr) in tools)
            Assert.False(
                string.IsNullOrWhiteSpace(attr.Title),
                $"Tool '{attr.Name}' in {type.Name}.{method.Name} should have a Title");
    }

    [Fact]
    public void AllTools_ShouldHaveUniqueName()
    {
        var tools = GetAllTools().ToList();
        var names = tools.Select(t => t.Attr.Name).ToList();
        var duplicates = names.GroupBy(n => n).Where(g => g.Count() > 1).Select(g => g.Key).ToList();

        Assert.Empty(duplicates);
    }

    [Fact]
    public void AllTools_TitleShouldBeHumanReadable()
    {
        var tools = GetAllTools().ToList();

        foreach (var (_, _, attr) in tools)
        {
            Assert.False(
                (attr.Title ?? "").Contains('_'),
                $"Tool '{attr.Name}' title should not contain underscores");

            Assert.True(
                char.IsUpper(attr.Title?[0] ?? ' '),
                $"Tool '{attr.Name}' title should start with uppercase");
        }
    }

    [Theory]
    [InlineData("convert_document", false, true, false)]
    public void ConversionTools_ShouldHaveCorrectAnnotations(
        string toolName,
        bool destructive,
        bool idempotent,
        bool readOnly)
    {
        var tool = GetAllTools().FirstOrDefault(t => t.Attr.Name == toolName);

        Assert.NotNull(tool.Attr);
        Assert.Equal(destructive, tool.Attr.Destructive);
        Assert.Equal(idempotent, tool.Attr.Idempotent);
        Assert.Equal(readOnly, tool.Attr.ReadOnly);
        Assert.False(tool.Attr.OpenWorld);
    }

    [Fact]
    public void SessionTool_ShouldBeDestructive()
    {
        var tool = GetAllTools().FirstOrDefault(t => t.Attr.Name == "document_session");

        Assert.NotNull(tool.Attr);
        Assert.True(tool.Attr.Destructive);
        Assert.False(tool.Attr.Idempotent);
        Assert.False(tool.Attr.ReadOnly);
    }

    [Fact]
    public void ReadOnlyTools_ShouldNotBeDestructive()
    {
        var readOnlyTools = GetAllTools().Where(t => t.Attr.ReadOnly).ToList();

        foreach (var (_, _, attr) in readOnlyTools)
            Assert.False(
                attr.Destructive,
                $"Tool '{attr.Name}' is ReadOnly but marked as Destructive - this is contradictory");
    }

    [Fact]
    public void ToolCount_ShouldMatchExpected()
    {
        var tools = GetAllTools().ToList();

        Assert.True(
            tools.Count >= 80,
            $"Expected at least 80 tools, but found {tools.Count}");
    }

    [Theory]
    [InlineData("word_")]
    [InlineData("excel_")]
    [InlineData("ppt_")]
    [InlineData("pdf_")]
    public void CategoryTools_ShouldExist(string prefix)
    {
        var categoryTools = GetAllTools().Where(t => t.Attr.Name?.StartsWith(prefix) == true).ToList();

        Assert.NotEmpty(categoryTools);
    }
}
