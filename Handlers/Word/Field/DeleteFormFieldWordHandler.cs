using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Handler for deleting form fields from Word documents.
/// </summary>
public class DeleteFormFieldWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_form_field";

    /// <summary>
    ///     Deletes one or more form fields from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: fieldName (single field), fieldNames (array of fields)
    ///     If neither provided, deletes all form fields.
    /// </param>
    /// <returns>Success message with deletion count.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteFormFieldParameters(parameters);

        var document = context.Document;
        var formFields = document.Range.FormFields;

        List<string> fieldsToDelete;
        if (p.FieldNames is { Length: > 0 })
            fieldsToDelete = p.FieldNames.Where(f => !string.IsNullOrEmpty(f)).ToList();
        else if (!string.IsNullOrEmpty(p.FieldName))
            fieldsToDelete = [p.FieldName];
        else
            fieldsToDelete = formFields.Select(f => f.Name).ToList();

        var deletedCount = 0;
        foreach (var name in fieldsToDelete)
        {
            var field = formFields[name];
            if (field != null)
            {
                field.Remove();
                deletedCount++;
            }
        }

        MarkModified(context);
        return Success($"Deleted {deletedCount} form field(s)");
    }

    /// <summary>
    ///     Extracts parameters for the delete form field operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DeleteFormFieldParameters ExtractDeleteFormFieldParameters(OperationParameters parameters)
    {
        var fieldName = parameters.GetOptional<string?>("fieldName");
        var fieldNames = parameters.GetOptional<string[]?>("fieldNames");

        return new DeleteFormFieldParameters(fieldName, fieldNames);
    }

    /// <summary>
    ///     Parameters for the delete form field operation.
    /// </summary>
    /// <param name="FieldName">The name of a single field to delete.</param>
    /// <param name="FieldNames">An array of field names to delete.</param>
    private sealed record DeleteFormFieldParameters(string? FieldName, string[]? FieldNames);
}
