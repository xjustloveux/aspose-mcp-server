namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Base class for PDF tool tests providing PDF-specific functionality
/// </summary>
public abstract class PdfTestBase : TestBase
{
    /// <summary>
    ///     Checks if Aspose.Pdf is running in evaluation mode.
    /// </summary>
    protected new static bool IsEvaluationMode(AsposeLibraryType libraryType = AsposeLibraryType.Pdf)
    {
        return TestBase.IsEvaluationMode(libraryType);
    }
}