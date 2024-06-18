using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word_Counter.UI;

namespace Word_Counter
{
    internal class Program
    {
        static void Main()
        {
            //WordCount.WordCount wc = new WordCount.WordCount();
            UI.UI ui = new UI.UI();
            string[] files = ui.GetFiles();

            foreach (string file in files)
            {
                Console.WriteLine(file);
            }

            Console.WriteLine("Press Enter to exit...");
            Console.ReadLine();
        }
    }
}
