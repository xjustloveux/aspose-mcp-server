namespace AsposeMcpServer.Core;

/// <summary>
///     Represents the result of loading Aspose licenses for enabled components.
/// </summary>
/// <param name="LoadedLicenses">The list of successfully loaded license names.</param>
/// <param name="EnabledComponents">The list of all enabled component names.</param>
public record LicenseLoadResult(List<string> LoadedLicenses, List<string> EnabledComponents);
