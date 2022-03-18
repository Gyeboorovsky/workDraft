using System;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using SaveOptions = SautinSoft.Document.SaveOptions;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            string filePath = @"..\..\..\example.docx";
            string changeFrom = "Mati";
            string changeTo = "Rafi";
            string font = "Arial";
            int size = 16;
            FindAndReplace(filePath, changeFrom, changeTo, font, size);
        }
        
        public static void FindAndReplace(string filePath, string from, string to, string font, int size)
        {
            DocumentCore dc = DocumentCore.Load(filePath);

            // Find "Bean" and Replace everywhere on "Joker"
            Regex regex = new Regex($"{from}", RegexOptions.IgnoreCase);

            // Start:

            foreach (ContentRange item in dc.Content.Find(regex).Reverse())
            {
                item.Replace($"{to}", new CharacterFormat()
                {
                    FontName = $"{font}", 
                    Size = size
                });
            }

            // Save the document as DOCX format.
            string savePath = Path.ChangeExtension(filePath, ".replaced.docx");
            dc.Save(savePath, SaveOptions.DocxDefault);

            // Open the original and result documents for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
        
        public static XDocument ExtractStylesPart(
            string filePath,
            bool getStylesWithEffectsPart = true)
        {
            // Retrieve the StylesWithEffects part. You could pass false in the 
            // second parameter to retrieve the Styles part instead.
            var styles = ExtractStylesPart(filePath, true);

            // If the part was retrieved, send the contents to the console.
            if (styles != null)
                Console.WriteLine(styles.ToString());

            return new XDocument();
        }
    }
}

