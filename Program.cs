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
            Application Word = new Application();

            string[] files = ui.GetFiles();

            foreach (string file in files)
            {
                Console.WriteLine(file);
            }

            processing.RemoveSymbols(files, Word);
            wordCount.Count(files, Word);

            Word.Quit();
            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }
    }
}