using System;
using System.Text.RegularExpressions;
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
                //Ask Word for Word Count (Which will be wrong, because it includes everything, including paragraph markers.
                int wordCount = doc.Words.Count;
                Console.WriteLine("Word count: " + wordCount);

                string textContent = doc.Content.Text.Trim();
                Console.WriteLine("Word Count 2: " + countWords(textContent));

                Document newDoc = Word.Documents.Add();

                // Insert the content
                newDoc.Content.Text = textContent;

                // Generate a new file path
                string newFilePath = file + "modified.docx";

                // Save the new document
                newDoc.SaveAs2(newFilePath);
                doc.Close();
            }
            Word.Quit();
        }

        private int countWords(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            // Split the text using a regex that matches words
            string[] words = Regex.Split(text, @"\W+");
            // Filter out empty entries and count the result
            return words.Length;
        }
    }
}