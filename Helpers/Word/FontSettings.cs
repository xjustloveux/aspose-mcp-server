namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Font settings for header/footer text.
/// </summary>
/// <param name="FontName">The font name.</param>
/// <param name="FontNameAscii">The ASCII font name.</param>
/// <param name="FontNameFarEast">The Far East font name.</param>
/// <param name="FontSize">The font size.</param>
public sealed record FontSettings(
    string? FontName,
    string? FontNameAscii,
    string? FontNameFarEast,
    double? FontSize);
