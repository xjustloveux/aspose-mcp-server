namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Represents the resolved command for starting an extension process.
/// </summary>
/// <param name="FileName">The executable file name to run.</param>
/// <param name="ArgumentList">The list of arguments to pass to the executable.</param>
public record CommandResolution(string FileName, List<string> ArgumentList);
