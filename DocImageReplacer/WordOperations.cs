using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocImageReplacer
{
    public static class WordOperations
    {
        public static void ReplaceTextWithImage(Microsoft.Office.Interop.Word.Document doc, string searchText, string imagePath)
        {
            Microsoft.Office.Interop.Word.Range range = doc.Content;
            range.Find.ClearFormatting();
            if (range.Find.Execute(searchText))
            {
                range.Text = "";
                object linkToFile = false;
                object saveWithDoc = true;
                doc.InlineShapes.AddPicture(imagePath, ref linkToFile, ref saveWithDoc, range);
            }
        }
    }
}
