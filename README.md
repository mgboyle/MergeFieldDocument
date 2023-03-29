# MergeFieldDocument
Here's an example C# console app that reads the input JSON file and a Word template document, and creates a new Word document with the merge fields replaced with the values from the JSON file:


To use the app, you need to pass the JSON file and the Word template file as command-line arguments. For example, if you have the data.json file and the template.docx file in the same directory as the app, you can run the following command:

```
WordMergeConsoleApp.exe data.json template.docx
```
The app will create a new Word document with the merge fields replaced with the values from data.json. The output file will have a filename in the format <template filename>-<refno>.docx, where <refno> is the value of the refno field in data.json.