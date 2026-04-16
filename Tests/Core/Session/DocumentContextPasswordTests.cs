using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Tests.Infrastructure;
using Document = Aspose.Words.Document;
using InvalidPasswordException = Aspose.Slides.InvalidPasswordException;
using OoxmlSaveOptions = Aspose.Words.Saving.OoxmlSaveOptions;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Regression tests for bug 20260415-session-pwd-swallow.
///     Fix lives in <see cref="DocumentContext{T}" /> private loaders
///     (<c>LoadWordDocument</c>, <c>LoadExcelWorkbook</c>, <c>LoadPowerPointPresentation</c>,
///     <c>LoadPdfDocument</c>).
///     Pre-fix: wrong password on an encrypted document was silently swallowed — the
///     loaders fell through to the no-password constructor and returned a "successfully"
///     opened document. For Excel the catch was additionally too broad
///     (<c>CellsException</c> root) and also masked non-password failures such as
///     <c>FileFormat</c> / <c>FileCorrupted</c>.
///     Post-fix: password-specific exceptions propagate on encrypted files; Excel
///     narrowed to <c>when (ex.Code == ExceptionType.IncorrectPassword)</c> so other
///     Cells failure codes are no longer swallowed.
/// </summary>
public class DocumentContextPasswordTests : TestBase
{
    private const string CorrectPassword = "correctPwd";
    private const string WrongPassword = "wrongPwd";

    // ---------------------------------------------------------------------
    // Fixture builders — generate small password-protected documents on the
    // fly using Aspose APIs so we never commit binary fixtures.
    // ---------------------------------------------------------------------

    private string CreateEncryptedWord(string fileName, string password)
    {
        var path = CreateTestFilePath(fileName);
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("protected content");
        var saveOptions = new OoxmlSaveOptions { Password = password };
        doc.Save(path, saveOptions);
        return path;
    }

    private string CreateEncryptedExcel(string fileName, string password)
    {
        var path = CreateTestFilePath(fileName);
        using var wb = new Workbook();
        wb.Worksheets[0].Cells["A1"].PutValue("protected");
        // Set the open-workbook password via WorkbookSettings.Password
        wb.Settings.Password = password;
        wb.Save(path);
        return path;
    }

    private string CreateEncryptedPresentation(string fileName, string password)
    {
        var path = CreateTestFilePath(fileName);
        using var pres = new Presentation();
        pres.ProtectionManager.Encrypt(password);
        pres.Save(path, SaveFormat.Pptx);
        return path;
    }

    private string CreateEncryptedPdf(string fileName, string password)
    {
        var path = CreateTestFilePath(fileName);
        using var pdf = new Aspose.Pdf.Document();
        pdf.Pages.Add();
        pdf.Encrypt(password, password, new Permissions(), CryptoAlgorithm.AESx128);
        pdf.Save(path);
        return path;
    }

    // ---------------------------------------------------------------------
    // Case 1 — Wrong password on encrypted file → must throw
    //   pre-fix: silent open via fallback constructor
    //   post-fix: propagates the Aspose-native password exception
    // ---------------------------------------------------------------------

    [Fact]
    public void Create_EncryptedWord_WithWrongPassword_ShouldThrowIncorrectPasswordException()
    {
        var path = CreateEncryptedWord("enc_word_wrong.docx", CorrectPassword);

        var ex = Record.Exception(() =>
            DocumentContext<Document>.Create(null, null, path, password: WrongPassword));

        Assert.NotNull(ex);
        Assert.IsType<IncorrectPasswordException>(ex);
    }

    [Fact]
    public void Create_EncryptedExcel_WithWrongPassword_ShouldThrowIncorrectPasswordCellsException()
    {
        var path = CreateEncryptedExcel("enc_excel_wrong.xlsx", CorrectPassword);

        var ex = Record.Exception(() =>
            DocumentContext<Workbook>.Create(null, null, path, password: WrongPassword));

        Assert.NotNull(ex);
        var cellsEx = Assert.IsType<CellsException>(ex);
        Assert.Equal(ExceptionType.IncorrectPassword, cellsEx.Code);
    }

    [Fact]
    public void Create_EncryptedPowerPoint_WithWrongPassword_ShouldThrowInvalidPasswordException()
    {
        var path = CreateEncryptedPresentation("enc_ppt_wrong.pptx", CorrectPassword);

        var ex = Record.Exception(() =>
            DocumentContext<Presentation>.Create(null, null, path, password: WrongPassword));

        Assert.NotNull(ex);
        Assert.IsType<InvalidPasswordException>(ex);
    }

    [Fact]
    public void Create_EncryptedPdf_WithWrongPassword_ShouldThrowInvalidPasswordException()
    {
        var path = CreateEncryptedPdf("enc_pdf_wrong.pdf", CorrectPassword);

        var ex = Record.Exception(() =>
            DocumentContext<Aspose.Pdf.Document>.Create(null, null, path, password: WrongPassword));

        Assert.NotNull(ex);
        Assert.IsType<Aspose.Pdf.InvalidPasswordException>(ex);
    }

    // ---------------------------------------------------------------------
    // Case 2 — Correct password on encrypted file → still works
    //   Guards against a naive fix that just rethrows for any catch path.
    // ---------------------------------------------------------------------

    [Fact]
    public void Create_EncryptedWord_WithCorrectPassword_ShouldLoadSuccessfully()
    {
        var path = CreateEncryptedWord("enc_word_ok.docx", CorrectPassword);

        using var context = DocumentContext<Document>.Create(null, null, path, password: CorrectPassword);

        Assert.NotNull(context.Document);
        Assert.False(context.IsSession);
        Assert.Equal(path, context.SourcePath);
    }

    [Fact]
    public void Create_EncryptedExcel_WithCorrectPassword_ShouldLoadSuccessfully()
    {
        var path = CreateEncryptedExcel("enc_excel_ok.xlsx", CorrectPassword);

        using var context = DocumentContext<Workbook>.Create(null, null, path, password: CorrectPassword);

        Assert.NotNull(context.Document);
        Assert.Equal("protected", context.Document.Worksheets[0].Cells["A1"].StringValue);
    }

    [Fact]
    public void Create_EncryptedPowerPoint_WithCorrectPassword_ShouldLoadSuccessfully()
    {
        var path = CreateEncryptedPresentation("enc_ppt_ok.pptx", CorrectPassword);

        using var context = DocumentContext<Presentation>.Create(null, null, path, password: CorrectPassword);

        Assert.NotNull(context.Document);
    }

    [Fact]
    public void Create_EncryptedPdf_WithCorrectPassword_ShouldLoadSuccessfully()
    {
        var path = CreateEncryptedPdf("enc_pdf_ok.pdf", CorrectPassword);

        using var context =
            DocumentContext<Aspose.Pdf.Document>.Create(null, null, path, password: CorrectPassword);

        Assert.NotNull(context.Document);
    }

    // ---------------------------------------------------------------------
    // Case 3 — Excel non-password CellsException (e.g. FileFormat / corrupt)
    //   must no longer be swallowed by the password catch.
    //   Pre-fix: catch (CellsException) matched the root hierarchy and
    //   silently fell through to `new Workbook(path)` which then raised a
    //   second, generic error — the original Code was lost.
    //   Post-fix: filter is `when (ex.Code == ExceptionType.IncorrectPassword)`
    //   so any other Code propagates directly on the password-path attempt.
    // ---------------------------------------------------------------------

    [Fact]
    public void Create_CorruptExcel_WithPassword_ShouldNotBeSwallowedByPasswordCatch()
    {
        // Garbage bytes at a .xlsx path — Aspose.Cells will raise a
        // CellsException whose Code is NOT IncorrectPassword (typically
        // FileFormat / InvalidData). Pre-fix this was swallowed.
        var path = CreateTestFilePath("corrupt.xlsx");
        File.WriteAllBytes(path, [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07]);

        var ex = Record.Exception(() =>
            DocumentContext<Workbook>.Create(null, null, path, password: "anyPwd"));

        Assert.NotNull(ex);
        // Must surface as a CellsException whose Code is specifically
        // NOT IncorrectPassword — i.e. the real root cause is preserved,
        // not masked as a password issue nor silently retried.
        var cellsEx = Assert.IsType<CellsException>(ex);
        Assert.NotEqual(ExceptionType.IncorrectPassword, cellsEx.Code);
    }

    // ---------------------------------------------------------------------
    // Case 4 — Regression guard: pre-existing "wrong password on an
    //   unprotected file" tests in DocumentContextTests.cs must still pass
    //   (impact.md confirms the catch is never entered for a clear file;
    //   the password is simply ignored by Aspose). We re-assert the same
    //   invariant here for each loader, colocated with the fix evidence.
    // ---------------------------------------------------------------------

    [Fact]
    public void Create_UnprotectedWord_WithAnyPassword_ShouldStillSucceed()
    {
        var path = CreateTestFilePath("plain_word.docx");
        new Document().Save(path);

        using var context = DocumentContext<Document>.Create(null, null, path, password: WrongPassword);

        Assert.NotNull(context.Document);
    }

    [Fact]
    public void Create_UnprotectedExcel_WithAnyPassword_ShouldStillSucceed()
    {
        var path = CreateTestFilePath("plain_excel.xlsx");
        using (var wb = new Workbook())
        {
            wb.Save(path);
        }

        using var context = DocumentContext<Workbook>.Create(null, null, path, password: WrongPassword);

        Assert.NotNull(context.Document);
    }

    [Fact]
    public void Create_UnprotectedPowerPoint_WithAnyPassword_ShouldStillSucceed()
    {
        var path = CreateTestFilePath("plain_ppt.pptx");
        using (var pres = new Presentation())
        {
            pres.Save(path, SaveFormat.Pptx);
        }

        using var context = DocumentContext<Presentation>.Create(null, null, path, password: WrongPassword);

        Assert.NotNull(context.Document);
    }

    [Fact]
    public void Create_UnprotectedPdf_WithAnyPassword_ShouldStillSucceed()
    {
        var path = CreateTestFilePath("plain_pdf.pdf");
        using (var pdf = new Aspose.Pdf.Document())
        {
            pdf.Pages.Add();
            pdf.Save(path);
        }

        using var context =
            DocumentContext<Aspose.Pdf.Document>.Create(null, null, path, password: WrongPassword);

        Assert.NotNull(context.Document);
    }
}
