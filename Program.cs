using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word_Counter.UI;
using Microsoft.Office.Interop.Word;
using System.Xml;
using System.IO;

namespace Word_Counter
{
    internal class Program
    {
        static void Main()
        {
            WordCount.WordCount wordCount = new WordCount.WordCount();
            UI.UI ui = new UI.UI();
            Processing.Processing processing = new Processing.Processing();
            Processing2.Processing2 processing2 = new Processing2.Processing2();
            Application Word = new Application();

            string[] files = ui.GetFiles();

            foreach (string file in files)
            {
                Console.WriteLine(file);
            }

            processing.RemoveSymbols(files, Word);
            files = Directory.GetFiles(ui.directoryPath, "*modified.docx");
            wordCount.Count(files, Word);

            Word.Quit();
            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }
    }
}