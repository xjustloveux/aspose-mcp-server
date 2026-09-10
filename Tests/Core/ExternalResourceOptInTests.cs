using AsposeMcpServer.Core;
using AsposeMcpServer.Tools.Conversion;

namespace AsposeMcpServer.Tests.Core;

/// <summary>
///     Covers A-03's opt-in half. Following references an MHT archive does not carry makes this
///     server issue outbound requests, so it is off by default. It used to be a per-request tool
///     parameter, which meant any caller could switch the policy off for their own request — on a
///     tool annotated <c>OpenWorld = false</c>. Whether the server may reach the network belongs
///     to the deployment, so it is now an operator setting.
/// </summary>
public class ExternalResourceOptInTests
{
    /// <summary>Runs an action with an environment variable set, restoring it afterwards.</summary>
    /// <param name="name">Variable name.</param>
    /// <param name="value">Value to set for the duration.</param>
    /// <param name="action">Action to run.</param>
    private static void WithEnvironmentVariable(string name, string? value, Action action)
    {
        var original = Environment.GetEnvironmentVariable(name);
        Environment.SetEnvironmentVariable(name, value);
        try
        {
            action();
        }
        finally
        {
            Environment.SetEnvironmentVariable(name, original);
        }
    }

    [Fact]
    public void ConvertTool_ShouldNotExposeTheSwitchToCallers()
    {
        var execute = typeof(ConvertDocumentTool).GetMethod(nameof(ConvertDocumentTool.Execute));
        Assert.NotNull(execute);

        var parameterNames = execute.GetParameters().Select(p => p.Name).ToList();

        Assert.DoesNotContain("allowExternalResources", parameterNames);
    }

    [Fact]
    public void Default_ShouldBeOff()
    {
        Assert.False(new ServerConfig().AllowExternalResources);
    }

    [Fact]
    public void CommandLineFlag_ShouldTurnItOn()
    {
        var config = ServerConfig.LoadFromArgs(["--allow-external-resources"]);

        Assert.True(config.AllowExternalResources);
    }

    [Fact]
    public void EnvironmentVariable_ShouldTurnItOn()
    {
        WithEnvironmentVariable("ASPOSE_ALLOW_EXTERNAL_RESOURCES", "true", () =>
        {
            var config = ServerConfig.LoadFromArgs([]);

            Assert.True(config.AllowExternalResources);
        });
    }

    [Fact]
    public void WithoutTheSetting_ShouldStayOff()
    {
        WithEnvironmentVariable("ASPOSE_ALLOW_EXTERNAL_RESOURCES", null, () =>
        {
            var config = ServerConfig.LoadFromArgs([]);

            Assert.False(config.AllowExternalResources);
        });
    }

    [Fact]
    public void TheSetterShouldNotBePublic()
    {
        // An operator setting that request handling could reassign would be no better than the
        // parameter it replaced.
        var property = typeof(ServerConfig).GetProperty(nameof(ServerConfig.AllowExternalResources));

        Assert.NotNull(property);
        Assert.False(property.SetMethod?.IsPublic ?? false);
    }
}
