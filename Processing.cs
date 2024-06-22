using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;


namespace Word_Counter.Processing
{
    public class Processing
    {
        public void removeSymbols(string[] files, Application Word)
        {
            foreach (string file in files)
            {
                Document doc = Word.Documents.Open(file);

                foreach(Range range in doc.StoryRanges)
                {
                    string text = range.Text;
                    range.Text = Regex.Replace(text, @"[.]", "");
                }
                doc.SaveAs(file + "modified.docx");
                doc.Close();
            }

        }
    }
}