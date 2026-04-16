using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;

namespace AsposeMcpServer.Errors.Ole;

/// <summary>
///     Translates raw Aspose / framework exceptions into sanitized BCL exceptions for
///     the OLE tool surface. Never stringifies inner-exception <see cref="Exception.Message" />
///     text into the emitted error — only fixed sentinels (from
///     <see cref="OleErrorMessageBuilder" />) or sanitized basenames flow through
///     (F-5, F-10). Preserves AC-19 cross-tool parity: every handler in the Word /
///     Excel / PowerPoint OLE tools uses this translator so the same failure mode surfaces
///     the same BCL type with the same sanitized message across the three tools.
/// </summary>
public static class OleErrorTranslator
{
    /// <summary>
    ///     Maps an arbitrary exception to a sanitized BCL exception. The mapping is:
    ///     password failures → <see cref="UnauthorizedAccessException" />;
    ///     index-out-of-range → <see cref="ArgumentOutOfRangeException" />;
    ///     permission / directory-missing → <see cref="UnauthorizedAccessException" />;
    ///     all other fallback cases → <see cref="InvalidOperationException" />.
    /// </summary>
    /// <param name="ex">The raw exception thrown from an Aspose API or IO call.</param>
    /// <param name="contextBasename">
    ///     Optional sanitized basename for contextual messages (e.g. the source filename
    ///     for save/remove failures). May be null.
    /// </param>
    /// <returns>A sanitized BCL exception ready to be thrown / logged.</returns>
    public static Exception Translate(Exception ex, string? contextBasename = null)
    {
        switch (ex)
        {
            case IncorrectPasswordException:
            case InvalidPasswordException:
                return new UnauthorizedAccessException(OleErrorMessageBuilder.InvalidPassword());

            // Aspose.Cells signals password failure via ExceptionType enum, not a dedicated exception type.
            case CellsException { Code: ExceptionType.IncorrectPassword }:
                return new UnauthorizedAccessException(OleErrorMessageBuilder.InvalidPassword());

            case ArgumentOutOfRangeException:
                return new ArgumentOutOfRangeException(
                    null, OleErrorMessageBuilder.IndexOutOfRange(null, null));

            case UnauthorizedAccessException:
            case DirectoryNotFoundException:
                return new UnauthorizedAccessException(
                    OleErrorMessageBuilder.OutputDirectoryNotWritable(contextBasename));

            case IOException:
                return new IOException(OleErrorMessageBuilder.SaveFailed(contextBasename));

            default:
                return new InvalidOperationException(OleErrorMessageBuilder.SaveFailed(contextBasename));
        }
    }
}
