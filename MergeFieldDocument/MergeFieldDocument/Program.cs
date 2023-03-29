using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using static System.Net.Mime.MediaTypeNames;

namespace WordMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            string jsonFile = args[0];
            string templateDocument = args[1];
            string outputDocument = args[2];

            Dictionary<string, string> mergeFields = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(jsonFile));

            using (WordprocessingDocument doc = WordprocessingDocument.Create(outputDocument, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                using (WordprocessingDocument templateDoc = WordprocessingDocument.Open(templateDocument, false))
                {
                    var docFields = templateDoc.MainDocumentPart.Document.Descendants<FieldCode>();
                    foreach (var field in docFields)
                    {
                        string fieldName = field.Text.Replace("MERGEFIELD ", "").Trim();
                        if (mergeFields.ContainsKey(fieldName))
                        {
                            string fieldValue = mergeFields[fieldName];
                            field.Parent.ReplaceChild(new SimpleField(new Run(new DocumentFormat.OpenXml.Drawing.Text(fieldValue))), field);
                        }
                    }
                }

                mainPart.Document.Save();
            }

            Console.WriteLine("Merge complete");
        }
    }
}
