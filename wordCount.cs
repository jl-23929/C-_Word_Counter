using System;
using Microsoft.Office.Interop.Word;

namespace Word_Counter.WordCount
{
    public class WordCount
    {
        public void ProcessDocs(string[] files)
        {
            Application Word = new Application();

            foreach (string file in files)
            {
                Document doc = Word.Documents.Open(file);
                int wordCount = doc.Words.Count;
                Console.WriteLine("Word count: " + wordCount);
                doc.Close();
            }
        }
    }
}