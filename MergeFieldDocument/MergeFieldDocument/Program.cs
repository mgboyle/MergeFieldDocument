using System;
using System.IO;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;

namespace WordMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            // Set the path to the JSON file and Word template document
            string jsonFilePath = args[0];
            string templateDocumentPath = args[1];

            // Read the JSON data from the file
            string json = File.ReadAllText(jsonFilePath);
            JObject data = JObject.Parse(json);

            // Create a new Word document
            Application word = new Application();
            Document doc = word.Documents.Add(templateDocumentPath);

            // Loop through the fields in the document and replace them with the data from the JSON
            foreach (Field field in doc.Fields)
            {
                string fieldName = field.Code.Text.Trim().Replace("MERGEFIELD", "").Trim();
                if (data[fieldName] != null)
                {
                    field.Select();
                    word.Selection.TypeText(data[fieldName].ToString());
                }
            }

            // Save the merged document and close Word
            string outputFilePath = Path.GetFileNameWithoutExtension(templateDocumentPath) + "_" + data["refno"].ToString() + ".docx";
            doc.SaveAs2(outputFilePath, WdSaveFormat.wdFormatDocumentDefault);
            doc.Close();
            word.Quit();
        }
    }
}
