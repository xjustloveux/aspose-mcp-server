using Aspose.Email;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Helper class for determining email save format options based on file extension.
/// </summary>
public static class EmailFormatHelper
{
    /// <summary>
    ///     Determines the appropriate <see cref="SaveOptions" /> for saving an email message
    ///     based on the output file extension.
    /// </summary>
    /// <param name="outputPath">The output file path whose extension determines the format.</param>
    /// <returns>The appropriate <see cref="SaveOptions" /> for the given file extension.</returns>
    public static SaveOptions DetermineEmailSaveFormat(string outputPath)
    {
        var ext = Path.GetExtension(outputPath).ToLowerInvariant();
        return ext switch
        {
            ".eml" => SaveOptions.DefaultEml,
            ".msg" => SaveOptions.DefaultMsgUnicode,
            ".mht" or ".mhtml" => SaveOptions.DefaultMhtml,
            ".html" or ".htm" => SaveOptions.DefaultHtml,
            _ => SaveOptions.DefaultEml
        };
    }
}
