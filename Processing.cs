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

                string[] symbols = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "=", "-", "{", "}", "[", "]", "|", "\\", ":", ";", "\"", "'", "<", ">", ",", ".", "?", "/", "`", "~" };

                foreach (Range range in doc.StoryRanges)
                {
                    foreach(string symbol in symbols)
                    {
                        Find findObject = range.Find;
                        findObject.ClearFormatting();
                        findObject.Text = symbol;
                        findObject.Replacement.ClearFormatting();
                        findObject.MatchWildcards = false;
                        findObject.Execute(Replace: WdReplace.wdReplaceAll);
                    }
                }
                doc.SaveAs(file + "modified.docx");
                doc.Close();
            }

        }
    }
}