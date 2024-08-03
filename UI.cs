using System;
using System.IO;

namespace Word_Counter.UI
{
    public class UI
    {

        public string directoryPath;
        public string[] GetFiles()
        {

            Console.Write("Enter Directory Path: ");
            directoryPath = Console.ReadLine();

            if (!Directory.Exists(directoryPath))
            {
                Console.WriteLine("Directory does not exist");
                throw new Exception("Directory does not exist");
            }
            Console.Write(directoryPath);
            string[] files = Directory.GetFiles(directoryPath, "*.docx");
            return files;
        }
    }
}