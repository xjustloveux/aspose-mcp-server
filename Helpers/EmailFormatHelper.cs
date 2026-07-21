using Aspose.Email;
using Aspose.Email.Tools;

namespace AsposeMcpServer.Helpers;

/// <summary>
///     Helper class for email format concerns: choosing save options for an output extension
///     and detecting the format of an existing file from its content.
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

    /// <summary>
    ///     Detects the email format name from the file content — not the extension, which the
    ///     file may carry incorrectly (e.g. EML content saved under a .msg name). When the content
    ///     is inconclusive (e.g. HTML exports, which are not email container formats), the
    ///     extension is used as a fallback.
    /// </summary>
    /// <param name="path">The email file path.</param>
    /// <returns>
    ///     The detected format name (e.g. "EML", "MSG", "MHTML", "HTML"), or "Unknown" when
    ///     neither the content nor the extension identifies the format.
    /// </returns>
    public static string DetectFormatName(string path)
    {
        try
        {
            var formatType = FileFormatUtil.DetectFileFormat(path).FileFormatType;
            if (formatType != FileFormatType.Unknown)
                return formatType == FileFormatType.Mht
                    ? "MHTML"
                    : formatType.ToString().ToUpperInvariant();
        }
        catch (Exception)
        {
            // Content probing failed — fall through to the extension fallback below.
        }

        return Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".eml" => "EML",
            ".msg" => "MSG",
            ".mhtml" or ".mht" => "MHTML",
            ".html" or ".htm" => "HTML",
            _ => "Unknown"
        };
    }
}
