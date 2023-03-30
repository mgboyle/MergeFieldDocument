using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;

namespace OpenXmlMergeJson
{
    class Program
    {
        static void Main(string[] args)
        {
            string jsonFilePath = args[0];
            string templateDocumentPath = args[1];
            string outputFilePath = Path.GetFileNameWithoutExtension(templateDocumentPath) + "_merged.docx";

            string json = File.ReadAllText(jsonFilePath);
            JObject data = JObject.Parse(json);

            File.Copy(templateDocumentPath, outputFilePath, true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(outputFilePath, true))
            {
                // Get the main document part
                MainDocumentPart mainPart = doc.MainDocumentPart;

                // Find all SimpleField elements in the document
                var simpleFields = mainPart.Document.Descendants<SimpleField>().ToList();

                // Replace the merge fields with the corresponding values from the JSON object
                foreach (var field in simpleFields)
                {
                    string fieldName = GetFieldName(field);

                    if (data[fieldName] != null)
                    {
                        Text text = new Text(data[fieldName].ToString());
                        Run run = new Run(new RunProperties(field.Run.RunProperties), text);
                        field.Parent.ReplaceChild(run, field);
                    }
                }

                // Save the changes to the document
                mainPart.Document.Save();
            }
        }

        private static string GetFieldName(SimpleField field)
        {
            string[] parts = field.Instruction.Value.Split(' ');

            if (parts.Length >= 3 && parts[0] == "MERGEFIELD")
            {
                return parts[1].Trim();
            }

            return null;
        }
    }
}
