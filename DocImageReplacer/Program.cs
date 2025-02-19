using DocImageReplacer;
using Microsoft.Office.Interop.Word;
using System;
using System.Configuration;
using System.IO;

class Program
{
    static void Main()
    {
        string docPath = ConfigurationManager.AppSettings["DocumentPath"];
        string imageFolder = ConfigurationManager.AppSettings["ImageFolder"];

        Application wordApp = new Application();
        Document doc = wordApp.Documents.Open(docPath);

        try
        {
            for (int i = 1; i <= 6; i++)
            {
                string pictureNumber = $"(picture {i})";
                string imagePath = Path.Combine(imageFolder, $"{i}.PNG");
                WordOperations.ReplaceTextWithImage(doc, pictureNumber, imagePath);
            }

            doc.Save();
            doc.Close();
            wordApp.Quit();
            Console.WriteLine("Images inserted successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}