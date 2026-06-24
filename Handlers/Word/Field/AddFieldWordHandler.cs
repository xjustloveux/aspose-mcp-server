using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for inserting fields in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Inserts a field at the specified paragraph position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldType (DATE, TIME, PAGE, NUMPAGES, AUTHOR, etc.)
    ///     Optional: fieldArgument, paragraphIndex, insertAtStart
    /// </param>
    /// <returns>Success message with field details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var fieldParams = ExtractFieldParameters(parameters);
        var document = context.Document;
        var builder = new DocumentBuilder(document);

        MoveToInsertPosition(builder, document, parameters, fieldParams.ParagraphIndex, fieldParams.InsertAtStart);

        var code = BuildFieldCode(fieldParams.FieldType, fieldParams.FieldArgument);
        var field = builder.InsertField(code);
        field.Update();

        MarkModified(context);

        return BuildResultMessage(fieldParams.FieldType, fieldParams.FieldArgument, code, field);
    }

    /// <summary>
    ///     Extracts and validates field parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted field parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when fieldType is not provided.</exception>
    private static FieldParameters ExtractFieldParameters(OperationParameters parameters)
    {
        var fieldType = parameters.GetOptional<string?>("fieldType");
        if (string.IsNullOrEmpty(fieldType))
            throw new ArgumentException("fieldType is required for add operation");

        return new FieldParameters(
            fieldType,
            parameters.GetOptional<string?>("fieldArgument"),
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional("insertAtStart", false)
        );
    }

    /// <summary>
    ///     Moves the document builder to the insertion position for the new field.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="document">The Word document.</param>
    /// <param name="parameters">The operation parameters (for resolving the paragraph address).</param>
    /// <param name="paragraphIndex">The target paragraph index (-1 or omitted = document end).</param>
    /// <param name="insertAtStart">Whether to insert at the start of the paragraph.</param>
    private static void MoveToInsertPosition(DocumentBuilder builder, Document document,
        OperationParameters parameters, int? paragraphIndex, bool insertAtStart)
    {
        if (!paragraphIndex.HasValue || paragraphIndex.Value == -1)
        {
            builder.MoveToDocumentEnd();
            return;
        }

        var targetPara = ParagraphResolver.Resolve(document, ParagraphAddress.From(parameters, paragraphIndex.Value))
            .Paragraph;

        // Anchor to the paragraph boundary, not to a Run: MoveTo(paragraph) lands at the paragraph
        // end (after any trailing field) and MoveTo(FirstChild) lands before any leading field, so
        // the inserted field becomes a sibling instead of nesting inside an existing field's range.
        builder.MoveTo(targetPara);
        if (insertAtStart && targetPara.FirstChild != null)
            builder.MoveTo(targetPara.FirstChild);
    }

    /// <summary>
    ///     Builds the field code from the field type and argument.
    /// </summary>
    /// <param name="fieldType">The field type.</param>
    /// <param name="fieldArgument">The optional field argument.</param>
    /// <returns>The constructed field code.</returns>
    private static string BuildFieldCode(string fieldType, string? fieldArgument)
    {
        var code = fieldType.ToUpper();
        return string.IsNullOrEmpty(fieldArgument) ? code : $"{code} {fieldArgument}";
    }

    /// <summary>
    ///     Builds the result message for the field insertion.
    /// </summary>
    /// <param name="fieldType">The field type.</param>
    /// <param name="fieldArgument">The field argument.</param>
    /// <param name="code">The field code.</param>
    /// <param name="field">The inserted field.</param>
    /// <returns>The result message.</returns>
    private static SuccessResult BuildResultMessage(string fieldType, string? fieldArgument, string code,
        Aspose.Words.Fields.Field field)
    {
        var message = $"Field inserted successfully\nField type: {fieldType}\n";
        if (!string.IsNullOrEmpty(fieldArgument))
            message += $"Field argument: {fieldArgument}\n";
        message += $"Field code: {code}\n";

        try
        {
            var fieldResult = field.Result;
            if (!string.IsNullOrEmpty(fieldResult))
                message += $"Field result: {fieldResult}";
        }
        catch
        {
            // Ignore errors reading field result (some fields may not have results)
        }

        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Record to hold field insertion parameters.
    /// </summary>
    /// <param name="FieldType">The field type.</param>
    /// <param name="FieldArgument">The field argument.</param>
    /// <param name="ParagraphIndex">The paragraph index.</param>
    /// <param name="InsertAtStart">Whether to insert at the start of the paragraph.</param>
    private sealed record FieldParameters(
        string FieldType,
        string? FieldArgument,
        int? ParagraphIndex,
        bool InsertAtStart);
}
