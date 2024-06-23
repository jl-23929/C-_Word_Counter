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

                RemoveBibliography(doc);
                
                    foreach (Paragraph paragraph in doc.Paragraphs)
                    {
                        string[] referenceTypes = { "*. [(][0-9]{4}[)].*", "*. [(][0-9]{4}[!)]@[)].*", "*. [(]n.d.[)].*" };

                        foreach (string reference in referenceTypes)
                        {

                            Find findObject = range.Find;
                            findObject.ClearFormatting();
                            findObject.Text = reference;
                            findObject.Replacement.ClearFormatting();
                            findObject.MatchWildcards = true;
                            findObject.Execute(Replace: WdReplace.wdReplaceAll);

                        }

                    }

                foreach (Range range in doc.StoryRanges)
                {
                    string[] inTextReferenceTypes = { "[(][!)]@, [0-9][0-9][0-9][0-9][)]", "[(][!)]@, [0-9][0-9][0-9][0-9]?[)]", "[(][!)]@, n.d.[)]" };

                    foreach (string reference in inTextReferenceTypes)
                    {

                        Find findObject = range.Find;
                        findObject.ClearFormatting();
                        findObject.Text = reference;
                        findObject.Replacement.ClearFormatting();
                        findObject.MatchWildcards = true;
                        findObject.Execute(Replace: WdReplace.wdReplaceAll);

                    }

                }


                foreach (Range range in doc.StoryRanges)
                {
                    string[] replace = { "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", ",", ".", "?", "!",
                    ":", ";", "(", ")", "[", "]", "{", "}", "/", "\\", "*", "+", "=", "|", "&", "^", "%",
                    "@", "~", "`", "'", "°", "𝜃", "×", "±", "≈", "∆", ">", "<", ">=", "<=", "=",
                    "ϕ", "φ", "Φ", "Ω" };

                    foreach (string symbol in replace)
                    {
                        Find findObject = range.Find;
                        findObject.ClearFormatting();
                        findObject.Text = symbol;
                        findObject.Replacement.ClearFormatting();
                        findObject.MatchWildcards = false;
                        findObject.Execute(Replace: WdReplace.wdReplaceAll);

                    }

                    Console.WriteLine("Story complete");
                }


                foreach (Range range in doc.StoryRanges)
                {
                    string[] remove = {"-", " M ", " V ", " Z ", " C ", " Q ", " Cu ", " Zn ", " Ag ",
                    " NO ", " KNO ", " MnO ", " NaCl ", " kPa ", " mL ", " L ", " aq ", " l ", " s ", " g ",
                    " x ", " KWh ", " kWh ", " cm ", " m ", " kW ", " W ", " MW ", " RPM ", " rpm ", " CO2 "};

                    foreach (string symbol in remove)
                    {
                        Find findObject = range.Find;
                        findObject.ClearFormatting();
                        findObject.Text = symbol;
                        findObject.Replacement.ClearFormatting();
                        findObject.MatchWildcards = false;
                        findObject.Execute(Replace: WdReplace.wdReplaceAll);
                    }

                    Console.WriteLine("Story complete");
                }


                doc.SaveAs(file + "modified.docx");
                doc.Close();
            }

        }

        public void RemoveBibliography(Document doc)
        {

            Range rng = doc.Content;
            Find findObject = rng.Find;
            findObject.Text = "Bibliography";
            findObject.MatchCase = true;
            findObject.MatchWholeWord = true;

            if (findObject.Execute())
            {
                // Move the start of the range to the end of the found text
                rng.Start = rng.End;
                // Move the end of the range to the end of the document
                rng.End = doc.Content.End;
                // Clear the range
                rng.Delete();

            }
        }
    }
}