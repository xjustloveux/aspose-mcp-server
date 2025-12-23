using System.Text;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for field operations in Word documents
///     Merges: WordInsertFieldTool, WordEditFieldTool, WordDeleteFieldTool, WordUpdateFieldTool,
///     WordGetFieldsTool, WordGetFieldDetailTool, WordAddFormFieldTool, WordEditFormFieldTool,
///     WordDeleteFormFieldTool, WordGetFormFieldsTool
/// </summary>
public class WordFieldTool : IAsposeTool
{
    public string Description =>
        @"Manage fields and form fields in Word documents. Supports 11 operations: insert_field, edit_field, delete_field, update_field, update_all, get_fields, get_field_detail, add_form_field, edit_form_field, delete_form_field, get_form_fields.

Usage examples:
- Insert field: word_field(operation='insert_field', path='doc.docx', fieldType='DATE', paragraphIndex=0)
- Edit field: word_field(operation='edit_field', path='doc.docx', fieldIndex=0, fieldArgument='yyyy-MM-dd')
- Delete field: word_field(operation='delete_field', path='doc.docx', fieldIndex=0)
- Update field: word_field(operation='update_field', path='doc.docx', fieldIndex=0)
- Update all fields: word_field(operation='update_all', path='doc.docx') or word_field(operation='update_field', path='doc.docx', updateAll=true)
- Get fields: word_field(operation='get_fields', path='doc.docx')
- Add form field: word_field(operation='add_form_field', path='doc.docx', formFieldType='TextInput', fieldName='name', paragraphIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'insert_field': Insert a field (required params: path, fieldType)
- 'edit_field': Edit a field (required params: path, fieldIndex)
- 'delete_field': Delete a field (required params: path, fieldIndex)
- 'update_field': Update a field (required params: path, fieldIndex optional if updateAll=true)
- 'update_all': Alias for update_field with updateAll=true (required params: path)
- 'get_fields': Get all fields (required params: path)
- 'get_field_detail': Get field details (required params: path, fieldIndex)
- 'add_form_field': Add a form field (required params: path, formFieldType, fieldName, options for DropDown type)
- 'edit_form_field': Edit a form field (required params: path, fieldName)
- 'delete_form_field': Delete a form field (required params: path, fieldName)
- 'get_form_fields': Get all form fields (required params: path)",
                @enum = new[]
                {
                    "insert_field", "edit_field", "delete_field", "update_field", "update_all", "get_fields",
                    "get_field_detail", "add_form_field", "edit_form_field", "delete_form_field", "get_form_fields"
                }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input, for write operations)"
            },
            // Field operations
            fieldType = new
            {
                type = "string",
                description =
                    "Field type: DATE, TIME, PAGE, NUMPAGES, AUTHOR, FILENAME, etc. (required for insert_field operation)",
                @enum = new[]
                {
                    "DATE", "TIME", "PAGE", "NUMPAGES", "AUTHOR", "FILENAME", "TITLE", "SUBJECT",
                    "COMPANY", "CREATEDATE", "SAVEDATE", "PRINTDATE", "USERNAME", "DOCPROPERTY",
                    "HYPERLINK", "REF", "MERGEFIELD", "IF", "SEQ"
                }
            },
            fieldArgument = new
            {
                type = "string",
                description = "Field argument (optional, for insert_field operation)"
            },
            paragraphIndex = new
            {
                type = "number",
                description =
                    "Paragraph index to insert into (0-based, optional, for insert_field operation). Valid range: 0 to (total paragraphs - 1), or -1 for document start. When not specified, inserts at document end."
            },
            insertAtStart = new
            {
                type = "boolean",
                description =
                    "Insert at the start of the paragraph (optional, default: false, for insert_field operation)"
            },
            fieldIndex = new
            {
                type = "number",
                description = "Field index (0-based, required for edit_field/delete_field/get_field_detail operations)"
            },
            fieldCode = new
            {
                type = "string",
                description = "New field code (optional, for edit_field operation)"
            },
            lockField = new
            {
                type = "boolean",
                description = "Lock the field (optional, for edit_field operation)"
            },
            unlockField = new
            {
                type = "boolean",
                description = "Unlock the field (optional, for edit_field operation)"
            },
            updateField = new
            {
                type = "boolean",
                description = "Update the field after editing (optional, default: true, for edit_field operation)"
            },
            keepResult = new
            {
                type = "boolean",
                description =
                    "Keep the field result text after deletion (optional, default: false, for delete_field operation)"
            },
            updateAll = new
            {
                type = "boolean",
                description =
                    "Update all fields (optional, default: false if fieldIndex provided, for update_field operation)"
            },
            includeCode = new
            {
                type = "boolean",
                description = "Include field code in results (optional, default: true, for get_fields operation)"
            },
            includeResult = new
            {
                type = "boolean",
                description = "Include field result in results (optional, default: true, for get_fields operation)"
            },
            // Form field operations
            formFieldType = new
            {
                type = "string",
                description = "Form field type: TextInput, CheckBox, DropDown (required for add_form_field operation)",
                @enum = new[] { "TextInput", "CheckBox", "DropDown" }
            },
            fieldName = new
            {
                type = "string",
                description = "Field name (required for form field operations)"
            },
            defaultValue = new
            {
                type = "string",
                description = "Default value (optional, for add_form_field/edit_form_field operations)"
            },
            options = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Options for dropdown (required for DropDown type, for add_form_field operation)"
            },
            checkedValue = new
            {
                type = "boolean",
                description =
                    "Checked state (optional, for CheckBox type, for add_form_field/edit_form_field operations)"
            },
            value = new
            {
                type = "string",
                description = "New value (optional, for TextInput type, for edit_form_field operation)"
            },
            selectedIndex = new
            {
                type = "number",
                description = "Selected option index (optional, for DropDown type, for edit_form_field operation)"
            },
            fieldNames = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of form field names to delete (optional, for delete_form_field operation)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        SecurityHelper.ValidateFilePath(path, allowAbsolutePaths: true);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        // Normalize operation name: update_all is an alias for update_field with updateAll=true
        if (operation == "update_all")
        {
            operation = "update_field";
            // Ensure updateAll is set to true if not already specified
            if (arguments != null && !arguments.ContainsKey("updateAll")) arguments["updateAll"] = true;
        }

        return operation switch
        {
            "insert_field" => await InsertFieldAsync(arguments, path, outputPath),
            "edit_field" => await EditFieldAsync(arguments, path, outputPath),
            "delete_field" => await DeleteFieldAsync(arguments, path, outputPath),
            "update_field" => await UpdateFieldAsync(arguments, path, outputPath),
            "get_fields" => await GetFieldsAsync(arguments, path),
            "get_field_detail" => await GetFieldDetailAsync(arguments, path),
            "add_form_field" => await AddFormFieldAsync(arguments, path, outputPath),
            "edit_form_field" => await EditFormFieldAsync(arguments, path, outputPath),
            "delete_form_field" => await DeleteFormFieldAsync(arguments, path, outputPath),
            "get_form_fields" => await GetFormFieldsAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Inserts a field into the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing fieldType, optional fieldArgument, paragraphIndex, runIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> InsertFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var fieldType = ArgumentHelper.GetString(arguments, "fieldType").ToUpper();
            var fieldArgument = ArgumentHelper.GetString(arguments, "fieldArgument", "");
            var paragraphIndex = ArgumentHelper.GetIntNullable(arguments, "paragraphIndex");
            var insertAtStart = ArgumentHelper.GetBool(arguments, "insertAtStart", false);

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);

            if (paragraphIndex.HasValue)
            {
                var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
                if (paragraphIndex.Value == -1)
                {
                    // paragraphIndex=-1 means document end, not document start
                    // Move to the last paragraph in the document body
                    var lastSection = doc.LastSection;
                    var bodyParagraphs = lastSection.Body.GetChildNodes(NodeType.Paragraph, false);
                    if (bodyParagraphs.Count > 0)
                    {
                        if (bodyParagraphs[^1] is Paragraph lastPara)
                        {
                            builder.MoveTo(lastPara);
                            if (insertAtStart)
                            {
                                // Move to start of last paragraph
                                if (lastPara.Runs.Count > 0) builder.MoveTo(lastPara.Runs[0]);
                            }
                            else
                            {
                                // Move to end of last paragraph
                                if (lastPara.Runs.Count > 0)
                                {
                                    var lastRun = lastPara.Runs[^1];
                                    builder.MoveTo(lastRun);
                                }
                            }
                        }
                        else
                        {
                            builder.MoveToDocumentEnd();
                        }
                    }
                    else
                    {
                        builder.MoveToDocumentEnd();
                    }
                }
                else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
                {
                    if (paragraphs[paragraphIndex.Value] is Paragraph targetPara)
                    {
                        if (insertAtStart)
                        {
                            builder.MoveTo(targetPara);
                            if (targetPara.Runs.Count > 0)
                                builder.MoveTo(targetPara.Runs[0]);
                            else
                                // Empty paragraph - just move to paragraph
                                builder.MoveTo(targetPara);
                        }
                        else
                        {
                            builder.MoveTo(targetPara);

                            if (targetPara.Runs.Count > 0)
                            {
                                var lastRun = targetPara.Runs[^1];
                                builder.MoveTo(lastRun);
                                // Move to end of the last run
                                try
                                {
                                    builder.MoveToParagraph(paragraphIndex.Value, lastRun.Text.Length);
                                }
                                catch (Exception ex)
                                {
                                    // If that fails, just stay at the run position
                                    // The field will be inserted after the run
                                    Console.Error.WriteLine(
                                        $"[WARN] Failed to move to paragraph {paragraphIndex.Value}, staying at run position: {ex.Message}");
                                }
                            }
                            else
                            {
                                // Empty paragraph - just move to paragraph
                                builder.MoveTo(targetPara);
                            }
                        }
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            $"Unable to find paragraph at index {paragraphIndex.Value}");
                    }
                }
                else
                {
                    throw new ArgumentException(
                        $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
                }
            }
            else
            {
                builder.MoveToDocumentEnd();
            }

            var fieldCode = fieldType;
            if (!string.IsNullOrEmpty(fieldArgument)) fieldCode += " " + fieldArgument;

            Field field;
            try
            {
                field = builder.InsertField(fieldCode);
                field.Update();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unable to insert field '{fieldCode}': {ex.Message}", ex);
            }

            doc.Save(outputPath);

            var result = "Field inserted successfully\n";
            result += $"Field type: {fieldType}\n";
            if (!string.IsNullOrEmpty(fieldArgument))
                result += $"Field argument: {fieldArgument}\n";
            result += $"Field code: {fieldCode}\n";

            try
            {
                var fieldResult = field.Result;
                if (!string.IsNullOrEmpty(fieldResult))
                    result += $"Field result: {fieldResult}\n";
            }
            catch (Exception ex)
            {
                // Field.Result may fail for some field types, but this is not critical
                Console.Error.WriteLine(
                    $"[WARN] Failed to get field result (may be normal for some field types): {ex.Message}");
                // Continue without the result information
            }

            if (paragraphIndex.HasValue)
            {
                if (paragraphIndex.Value == -1)
                    result += "Insert position: beginning of document\n";
                else
                    result += $"Insert position: paragraph #{paragraphIndex.Value}\n";
            }
            else
            {
                result += "Insert position: end of document\n";
            }

            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Edits an existing field
    /// </summary>
    /// <param name="arguments">JSON arguments containing fieldIndex, optional fieldType, fieldArgument</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> EditFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var fieldIndex = ArgumentHelper.GetInt(arguments, "fieldIndex");
            var fieldCode = ArgumentHelper.GetStringNullable(arguments, "fieldCode");
            var lockField = ArgumentHelper.GetBoolNullable(arguments, "lockField");
            var unlockField = ArgumentHelper.GetBoolNullable(arguments, "unlockField");
            var updateField = ArgumentHelper.GetBool(arguments, "updateField", true);

            var doc = new Document(path);
            var fields = doc.Range.Fields.ToList();

            if (fieldIndex < 0 || fieldIndex >= fields.Count)
                throw new ArgumentException(
                    $"Field index {fieldIndex} is out of range (document has {fields.Count} fields)");

            var field = fields[fieldIndex];
            var oldFieldCode = field.GetFieldCode();
            var oldResult = field.Result ?? "";
            var oldLocked = field.IsLocked;

            var changes = new List<string>();

            if (!string.IsNullOrEmpty(fieldCode))
                try
                {
                    var fieldStart = field.Start;
                    var fieldEnd = field.End;

                    if (fieldStart != null && fieldEnd != null)
                    {
                        var builder = new DocumentBuilder(doc);
                        builder.MoveTo(fieldStart);

                        var currentNode = fieldStart.NextSibling;
                        while (currentNode != null && currentNode != fieldEnd)
                        {
                            var nextNode = currentNode.NextSibling;
                            if (currentNode.NodeType != NodeType.FieldSeparator &&
                                currentNode.NodeType != NodeType.FieldEnd) currentNode.Remove();
                            currentNode = nextNode;
                        }

                        builder.MoveTo(fieldStart);
                        builder.Write(fieldCode);

                        changes.Add($"Field code updated: {oldFieldCode} -> {fieldCode}");
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Unable to update field code: {ex.Message}", ex);
                }

            if (lockField.HasValue && lockField.Value)
            {
                field.IsLocked = true;
                changes.Add("Field locked");
            }
            else if (unlockField.HasValue && unlockField.Value)
            {
                field.IsLocked = false;
                changes.Add("Field unlocked");
            }

            if (updateField)
            {
                field.Update();
                doc.UpdateFields();
            }

            doc.Save(outputPath);

            var result = $"Field #{fieldIndex} edited successfully\n";
            result += $"Original field code: {oldFieldCode}\n";
            if (!string.IsNullOrEmpty(fieldCode)) result += $"New field code: {fieldCode}\n";
            result += $"Original result: {oldResult}\n";
            result += $"Original lock status: {(oldLocked ? "Locked" : "Unlocked")}\n";
            if (changes.Count > 0) result += $"Changes: {string.Join(", ", changes)}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes a field from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing fieldIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var fieldIndex = ArgumentHelper.GetInt(arguments, "fieldIndex");
            var keepResult = ArgumentHelper.GetBool(arguments, "keepResult", false);

            var doc = new Document(path);
            var fields = doc.Range.Fields.ToList();

            if (fieldIndex < 0 || fieldIndex >= fields.Count)
                throw new ArgumentException(
                    $"Field index {fieldIndex} is out of range (document has {fields.Count} fields)");

            var field = fields[fieldIndex];
            var fieldType = field.Type.ToString();
            var fieldCode = field.GetFieldCode();
            var fieldResult = field.Result ?? "";

            try
            {
                if (keepResult)
                {
                    var fieldStart = field.Start;
                    var fieldEnd = field.End;

                    if (fieldStart != null && fieldEnd != null)
                    {
                        var nodesToRemove = new List<Node>();
                        var currentNode = (Node)fieldStart;
                        var endNode = (Node)fieldEnd;

                        while (currentNode != null && currentNode != endNode.NextSibling)
                        {
                            if (currentNode.NodeType == NodeType.FieldStart ||
                                currentNode.NodeType == NodeType.FieldSeparator)
                                nodesToRemove.Add(currentNode);
                            currentNode = currentNode.NextSibling;

                            if (currentNode == endNode) break;
                        }

                        foreach (var node in nodesToRemove) node.Remove();

                        endNode.Remove();
                    }
                }
                else
                {
                    var fieldStart = field.Start;
                    var fieldEnd = field.End;

                    if (fieldStart != null && fieldEnd != null)
                    {
                        var nodesToRemove = new List<Node>();
                        var currentNode = (Node)fieldStart;
                        var endNode = (Node)fieldEnd;

                        while (currentNode != null)
                        {
                            var nextNode = currentNode.NextSibling;
                            nodesToRemove.Add(currentNode);

                            if (currentNode == endNode) break;

                            currentNode = nextNode;
                        }

                        foreach (var node in nodesToRemove) node.Remove();
                    }
                }

                doc.Save(outputPath);

                var remainingFields = doc.Range.Fields.Count;

                var result = $"Field #{fieldIndex} deleted successfully\n";
                result += $"Type: {fieldType}\n";
                result += $"Code: {fieldCode}\n";
                if (!string.IsNullOrEmpty(fieldResult)) result += $"Result: {fieldResult}\n";
                result += $"Keep result text: {(keepResult ? "Yes" : "No")}\n";
                result += $"Remaining fields in document: {remainingFields}\n";
                result += $"Output: {outputPath}";

                return result;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Unable to delete field: {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    ///     Updates field values
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional fieldIndex (if null, updates all)</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message with update count</returns>
    private Task<string> UpdateFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var fieldIndex = ArgumentHelper.GetIntNullable(arguments, "fieldIndex");
            var fieldTypeFilter = ArgumentHelper.GetStringNullable(arguments, "fieldType");
            var updateAll = ArgumentHelper.GetBool(arguments, "updateAll", !fieldIndex.HasValue);

            var doc = new Document(path);
            var fields = doc.Range.Fields.ToList();
            var updatedCount = 0;
            var errors = new List<string>();

            if (fieldIndex.HasValue && !updateAll)
            {
                if (fieldIndex.Value < 0 || fieldIndex.Value >= fields.Count)
                    throw new ArgumentException(
                        $"Field index {fieldIndex.Value} is out of range (document has {fields.Count} fields)");

                var field = fields[fieldIndex.Value];
                try
                {
                    var oldResult = field.Result ?? "";
                    field.Update();
                    var newResult = field.Result ?? "";

                    doc.Save(outputPath);

                    var singleResult = $"Field #{fieldIndex.Value} updated successfully\n";
                    singleResult += $"Type: {field.Type}\n";
                    singleResult += $"Code: {field.GetFieldCode()}\n";
                    singleResult += $"Old result: {oldResult}\n";
                    singleResult += $"New result: {newResult}\n";
                    singleResult += $"Output: {outputPath}";

                    return singleResult;
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Unable to update field #{fieldIndex.Value}: {ex.Message}",
                        ex);
                }
            }

            foreach (var field in fields)
            {
                if (!string.IsNullOrEmpty(fieldTypeFilter))
                {
                    var filterType = fieldTypeFilter.ToUpper();
                    var fieldTypeName = field.Type.ToString().ToUpper();
                    if (!fieldTypeName.Contains(filterType)) continue;
                }

                try
                {
                    field.Update();
                    updatedCount++;
                }
                catch (Exception ex)
                {
                    errors.Add($"Field {field.Type} (index {fields.IndexOf(field)}): {ex.Message}");
                }
            }

            doc.Save(outputPath);

            var result = $"Successfully updated {updatedCount} fields\n";
            if (!string.IsNullOrEmpty(fieldTypeFilter)) result += $"Filter type: {fieldTypeFilter}\n";

            if (errors.Count > 0)
            {
                result += $"\nFailed to update fields ({errors.Count}):\n";
                foreach (var error in errors) result += $"  - {error}\n";
            }

            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Gets all fields from the document
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all fields</returns>
    private Task<string> GetFieldsAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var fieldTypeFilter = ArgumentHelper.GetStringNullable(arguments, "fieldType");
            var includeCode = ArgumentHelper.GetBool(arguments, "includeCode", true);
            var includeResult = ArgumentHelper.GetBool(arguments, "includeResult", true);

            var doc = new Document(path);
            var fields = new List<FieldInfo>();
            var fieldIndex = 0;

            foreach (var field in doc.Range.Fields)
            {
                if (!string.IsNullOrEmpty(fieldTypeFilter))
                {
                    var filterType = fieldTypeFilter.ToUpper();
                    var fieldTypeName = field.Type.ToString().ToUpper();
                    if (!fieldTypeName.Contains(filterType) && filterType != "ALL") continue;
                }

                var fieldInfo = new FieldInfo
                {
                    Index = fieldIndex++,
                    Type = field.Type.ToString(),
                    Code = field.GetFieldCode(),
                    Result = includeResult ? field.Result ?? "" : null,
                    IsLocked = field.IsLocked,
                    IsDirty = field.IsDirty
                };

                if (field is FieldHyperlink hyperlinkField)
                    fieldInfo.ExtraInfo =
                        $"Address: {hyperlinkField.Address ?? ""}, ScreenTip: {hyperlinkField.ScreenTip ?? ""}";
                else if (field is FieldRef refField) fieldInfo.ExtraInfo = $"Bookmark: {refField.BookmarkName ?? ""}";

                fields.Add(fieldInfo);
            }

            var result = new StringBuilder();
            result.AppendLine($"Document has {fields.Count} fields\n");

            if (fields.Count == 0)
            {
                result.AppendLine("No fields found");
                return result.ToString();
            }

            result.AppendLine("[Field List]");
            result.AppendLine(new string('-', 80));

            foreach (var fieldInfo in fields)
            {
                result.AppendLine($"Index: {fieldInfo.Index}");
                result.AppendLine($"Type: {fieldInfo.Type}");

                if (includeCode) result.AppendLine($"Code: {fieldInfo.Code}");

                if (includeResult && fieldInfo.Result != null) result.AppendLine($"Result: {fieldInfo.Result}");

                if (fieldInfo.ExtraInfo != null) result.AppendLine($"Extra info: {fieldInfo.ExtraInfo}");

                result.AppendLine($"Locked: {(fieldInfo.IsLocked ? "Yes" : "No")}");
                result.AppendLine($"Needs update: {(fieldInfo.IsDirty ? "Yes" : "No")}");
                result.AppendLine(new string('-', 80));
            }

            var typeGroups = fields.GroupBy(f => f.Type).OrderBy(g => g.Key);
            result.AppendLine("\n[Statistics by Type]");
            foreach (var group in typeGroups) result.AppendLine($"{group.Key}: {group.Count()}");

            return result.ToString();
        });
    }

    /// <summary>
    ///     Gets detailed information about a specific field
    /// </summary>
    /// <param name="arguments">JSON arguments containing fieldIndex</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with field details</returns>
    private Task<string> GetFieldDetailAsync(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var fieldIndex = ArgumentHelper.GetInt(arguments, "fieldIndex");

            var doc = new Document(path);
            var fields = doc.Range.Fields.ToList();

            if (fieldIndex < 0 || fieldIndex >= fields.Count)
                throw new ArgumentException(
                    $"Field index {fieldIndex} is out of range (document has {fields.Count} fields)");

            var field = fields[fieldIndex];
            var result = new StringBuilder();

            result.AppendLine("[Field Details]");
            result.AppendLine(new string('=', 80));
            result.AppendLine($"Index: {fieldIndex}");
            result.AppendLine($"Type: {field.Type}");
            result.AppendLine($"Type code: {(int)field.Type}");
            result.AppendLine($"Code: {field.GetFieldCode()}");
            result.AppendLine($"Result: {field.Result ?? "(No result)"}");
            result.AppendLine($"Locked: {(field.IsLocked ? "Yes" : "No")}");
            result.AppendLine($"Needs update: {(field.IsDirty ? "Yes" : "No")}");

            if (field is FieldHyperlink hyperlinkField)
            {
                result.AppendLine($"Address: {hyperlinkField.Address ?? "(None)"}");
                result.AppendLine($"Screen tip: {hyperlinkField.ScreenTip ?? "(None)"}");
            }
            else if (field is FieldRef refField)
            {
                result.AppendLine($"Bookmark name: {refField.BookmarkName ?? "(None)"}");
            }

            result.AppendLine(new string('=', 80));

            return result.ToString();
        });
    }

    /// <summary>
    ///     Adds a form field to the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing fieldType, name, optional defaultValue, paragraphIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> AddFormFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var fieldType = ArgumentHelper.GetString(arguments, "formFieldType");
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
            var defaultValue = ArgumentHelper.GetStringNullable(arguments, "defaultValue");
            var checkedValue = ArgumentHelper.GetBoolNullable(arguments, "checkedValue");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            switch (fieldType.ToLower())
            {
                case "textinput":
                    builder.InsertTextInput(fieldName, TextFormFieldType.Regular, "", defaultValue ?? "", 0);
                    break;

                case "checkbox":
                    builder.InsertCheckBox(fieldName, checkedValue ?? false, 0);
                    break;

                case "dropdown":
                    var optionsArray = ArgumentHelper.GetArray(arguments, "options");
                    if (optionsArray.Count == 0)
                        throw new ArgumentException("options array cannot be empty for DropDown type");
                    var options = optionsArray.Select(o => o?.GetValue<string>()).Where(o => !string.IsNullOrEmpty(o))
                        .ToArray();
                    builder.InsertComboBox(fieldName, options, 0);
                    break;

                default:
                    throw new ArgumentException(
                        $"Invalid fieldType: {fieldType}. Must be 'TextInput', 'CheckBox', or 'DropDown'");
            }

            doc.Save(outputPath);
            return $"{fieldType} field '{fieldName}' added: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an existing form field
    /// </summary>
    /// <param name="arguments">JSON arguments containing fieldIndex, optional name, defaultValue</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> EditFormFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
            var value = ArgumentHelper.GetStringNullable(arguments, "value");
            var checkedValue = ArgumentHelper.GetBoolNullable(arguments, "checkedValue");
            var selectedIndex = ArgumentHelper.GetIntNullable(arguments, "selectedIndex");

            var doc = new Document(path);
            var field = doc.Range.FormFields[fieldName];

            if (field == null) throw new ArgumentException($"Form field '{fieldName}' not found");

            if (field.Type == FieldType.FieldFormTextInput && value != null)
                field.Result = value;
            else if (field.Type == FieldType.FieldFormCheckBox && checkedValue.HasValue)
                field.Checked = checkedValue.Value;
            else if (field.Type == FieldType.FieldFormDropDown && selectedIndex.HasValue)
                if (selectedIndex.Value >= 0 && selectedIndex.Value < field.DropDownItems.Count)
                    field.DropDownSelectedIndex = selectedIndex.Value;

            doc.Save(outputPath);
            return $"Form field '{fieldName}' updated: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a form field from the document
    /// </summary>
    /// <param name="arguments">JSON arguments containing fieldIndex</param>
    /// <param name="path">Word document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteFormFieldAsync(JsonObject? arguments, string path, string outputPath)
    {
        return Task.Run(() =>
        {
            var fieldName = ArgumentHelper.GetStringNullable(arguments, "fieldName");
            var fieldNamesArray = ArgumentHelper.GetArray(arguments, "fieldNames", false);

            var doc = new Document(path);
            var formFields = doc.Range.FormFields;

            List<string> fieldsToDelete;
            if (fieldNamesArray is { Count: > 0 })
                fieldsToDelete = fieldNamesArray.Select(f => f?.GetValue<string>()).Where(f => !string.IsNullOrEmpty(f))
                    .Select(f => f!).ToList();
            else if (!string.IsNullOrEmpty(fieldName))
                fieldsToDelete = [fieldName];
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

            doc.Save(outputPath);
            return $"Deleted {deletedCount} form field(s): {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all form fields from the document
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Word document file path</param>
    /// <returns>Formatted string with all form fields</returns>
    private Task<string> GetFormFieldsAsync(JsonObject? _, string path)
    {
        return Task.Run(() =>
        {
            var doc = new Document(path);
            var sb = new StringBuilder();

            sb.AppendLine("=== Form Fields ===");
            sb.AppendLine();

            var formFields = doc.Range.FormFields.ToList();
            for (var i = 0; i < formFields.Count; i++)
            {
                var field = formFields[i];
                sb.AppendLine($"[{i + 1}] Name: {field.Name}");
                sb.AppendLine($"    Type: {field.Type}");

                switch (field.Type)
                {
                    case FieldType.FieldFormTextInput:
                        sb.AppendLine($"    Value: {field.Result ?? "(empty)"}");
                        break;
                    case FieldType.FieldFormCheckBox:
                        sb.AppendLine($"    Checked: {field.Checked}");
                        break;
                    case FieldType.FieldFormDropDown:
                        sb.AppendLine($"    Selected Index: {field.DropDownSelectedIndex}");
                        sb.AppendLine($"    Options: {string.Join(", ", field.DropDownItems)}");
                        break;
                }

                sb.AppendLine();
            }

            sb.AppendLine($"Total Form Fields: {formFields.Count}");

            return sb.ToString();
        });
    }

    private class FieldInfo
    {
        public int Index { get; init; }
        public string Type { get; init; } = "";
        public string Code { get; init; } = "";
        public string? Result { get; init; }
        public bool IsLocked { get; init; }
        public bool IsDirty { get; init; }
        public string? ExtraInfo { get; set; }
    }
}