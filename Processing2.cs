using System;
using OfficeIMO.Word;
using System.Text.RegularExpressions;

namespace Word_Counter.Processing2
{
    public class Processing2
    {
        public void RemoveSymbols(string[] files)
        {
            foreach (string file in files)
            {
                using (var document = WordDocument.Load(file))
                {
                    // Define the regex pattern to find and replace
                    string pattern = @" ";
                    string replacement = "";
                    Regex regex = new Regex(pattern);

                    // Iterate through paragraphs and perform regex replacement
                    foreach (var paragraph in document.Paragraphs)
                    {
                        if (!string.IsNullOrEmpty(paragraph.Text))
                        {
                            paragraph.Text = regex.Replace(paragraph.Text, replacement);
                        }
                    }

                    // Save the document
                    document.Save(file + "_modified.docx");
                }
            }
        }
    }
}