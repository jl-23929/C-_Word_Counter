using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace Word_Counter.WordCount
{
    public class WordCount
    {
        public void Count(string[] files, Application Word)
        {
            //Creates a new instance of Word.

            foreach (string file in files)
            {
                //Opens the document.
                Document doc = Word.Documents.Open(file);
                //Asks Word for the correct word count.
                int wordCount = doc.ComputeStatistics(WdStatistic.wdStatisticWords, IncludeFootnotesAndEndnotes: true);

                //Prints the word count.
                Console.WriteLine("Word Count 2: " + wordCount);

                //Closese the doc without saving it.
                doc.Close(false);
            }
            Word.Quit();
        }
    }
}