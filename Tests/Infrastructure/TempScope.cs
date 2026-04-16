namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Disposable wrapper around a temporary directory created by <see cref="SymlinkFixture" />.
///     Deletes the directory and all of its contents on disposal.
/// </summary>
public sealed class TempScope : IDisposable
{
    /// <summary>
    ///     Initialises a new <see cref="TempScope" /> backed by <paramref name="root" />.
    /// </summary>
    /// <param name="root">The absolute path of the temporary directory to manage.</param>
    public TempScope(string root)
    {
        Root = root;
    }

    /// <summary>The absolute path of the managed temporary directory.</summary>
    public string Root { get; }

    /// <inheritdoc />
    public void Dispose()
    {
        try
        {
            if (Directory.Exists(Root))
                Directory.Delete(Root, true);
        }
        catch
        {
            // Best-effort cleanup; test isolation does not depend on perfect teardown.
        }
    }
}
