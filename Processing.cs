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
                    Find findObject = range.Find;
                    findObject.ClearFormatting();
                    findObject.Text = ".";
                    findObject.Replacement.ClearFormatting();
                    findObject.MatchWildcards = false;
                    findObject.Execute(Replace: WdReplace.wdReplaceAll);
                }
                doc.SaveAs(file + "modified.docx");
                doc.Close();
            }

        }
    }
}