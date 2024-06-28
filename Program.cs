using System;
using Microsoft.Office.Interop.Word;

class Program
{
    static void Main()
    {
        Application wordApp = new Application();
        Document doc = null;

        try
        {
            // Open the Word document
            object fileName = @"C:\path\to\your\document.docx";
            object missing = Type.Missing;
            doc = wordApp.Documents.Open(ref fileName, ref missing, ref missing, ref missing);

            
            string[] commonReferencePatterns = { "(", ")" };

            foreach (Paragraph paragraph in doc.Paragraphs)
            {
                string text = paragraph.Range.Text;

                // Check if the paragraph contains a potential reference
                if (ContainsAPAReference(text, commonReferencePatterns))
                {
                    // If found, delete the paragraph
                    paragraph.Range.Delete();
                }
            }

            // Save the modified document
            doc.SaveAs2(@"C:\path\to\your\modified_document.docx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            if (doc != null)
            {
                doc.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            }
            wordApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
        }
    }

    static bool ContainsAPAReference(string text, string[] patterns)
    {
        foreach (string pattern in patterns)
        {
            if (text.Contains(pattern))
            {
                return true;
            }
        }
        return false;
    }
}