using DocumentFormat.OpenXml.Wordprocessing;

namespace KJBConseil.WordPubliposting;

internal static class ReplaceVariableByValues
{
    /// <summary>
    /// Expect an OpenXML document with configured fields.
    /// </summary>
    /// <param name="body"></param>
    /// <param name="fieldsToUpdate">Key must be the name of the field used in you document; value must be the replaced value.</param>
    /// <remarks>Fields are available in the "QuickPart" menu of an Office Document.
    /// You can choose to display, or hide, them with ALT+F9 key combination.
    /// CTRL+F9 will enable you to add them quickly where you cursor will be on the document.
    /// </remarks>
    /// <exception cref="InvalidOperationException"></exception>
    public static void Execute(Body body, Dictionary<string, string> fieldsToUpdate)
    {
        foreach (var fieldToUpdate in fieldsToUpdate)
        {
            var fieldName = fieldToUpdate.Key;
            var fieldNewValue = fieldToUpdate.Value;

            foreach (var parent in body.Descendants<FieldCode>()
                .Where(fieldCode => fieldCode.Text.Contains($"DOCVARIABLE  {fieldName}") || fieldCode.Text.Contains($"{fieldName}"))
                .Select(matchedFieldCode => matchedFieldCode.Parent))
            {
                if (parent is null)
                {
                    throw new InvalidOperationException($"The parent of the found '{fieldName}' field code should not be null.");
                }

                // to remove the doc variable declaration and replace by the targeted value
                parent.RemoveAllChildren<FieldCode>();
                parent.AppendChild(new Text(fieldNewValue));

                // search and delete opening curly bracket when using doc variable (CTRL+F9)
                var parentThatCouldHoldAnOpeningCurlyBracket =
                    parent.ElementsBefore().FirstOrDefault(before => before.Descendants<FieldChar>().FirstOrDefault() != null);
                var couldBeAnOpeningCurlyBracket = parentThatCouldHoldAnOpeningCurlyBracket?.GetFirstChild<FieldChar>();
                if (couldBeAnOpeningCurlyBracket != null &&
                    couldBeAnOpeningCurlyBracket.FieldCharType?.Value == FieldCharValues.Begin)
                {
                    parentThatCouldHoldAnOpeningCurlyBracket!.Remove();
                }

                // search and delete closing curly bracket when using doc variable (CTRL+F9)
                var parentThatCouldHoldAClosingCurlyBracket =
                    parent.ElementsAfter().FirstOrDefault(before => before.Descendants<FieldChar>().FirstOrDefault() != null);
                var couldBeAClosingCurlyBracket = parentThatCouldHoldAClosingCurlyBracket?.GetFirstChild<FieldChar>();
                if (couldBeAClosingCurlyBracket != null &&
                    couldBeAClosingCurlyBracket.FieldCharType?.Value == FieldCharValues.End)
                {
                    parentThatCouldHoldAClosingCurlyBracket!.Remove();
                }
            }
        }
    }
}
