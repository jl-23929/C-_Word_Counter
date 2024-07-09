using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace Word_Counter.Processing
{
    public class Processing
    {
        public void RemoveSymbols(string[] files, Application wordApp)
        {

            string[] inTextReferenceTypes = {
                        "[(][!)]@, [0-9][0-9][0-9][0-9][)]",
                        "[(][!)]@, [0-9][0-9][0-9][0-9]?[)]",
                        "[(][!)]@, n.d.[)]"
                    };

            string[] replaceSymbols = {
                        "1", "2", "3", "4", "5", "6", "7", "8", "9", "0",
                        ",", ".", "?", "!", ":", ";", "(", ")", "[", "]", "{", "}",
                        "/", "\\", "*", "+", "=", "|", "&", "^", "%", "@", "~",
                        "`", "'", "°", "𝜃", "×", "±", "≈", "∆", ">", "<", ">=",
                        "<=", "=", "ϕ", "φ", "Φ", "Ω", "Ω", "∑", "∞", "√"
                    };

            string[] removeSymbols = {
                        "-", " M ", " V ", " Z ", " C ", " Q ", " Cu ", " Zn ",
                        " Ag ", " NO ", " KNO ", " MnO ", " NaCl ", " kPa ",
                        " mL ", " L ", " aq ", " l ", " s ", " g ", " x ", " KWh ",
                        " kWh ", " cm ", " m ", " kW ", " W ", " MW ", " RPM ",
                        " rpm ", " CO2 "
                    };

            string[] referenceTypes = { @"^13[!^13]@ [(][0-9]{4}[)].[!^13]@^13", @"^13[!^13]@ [(][0-9]{4}[!)]@[)].[!^13]@^13" };

            foreach (string file in files)            
            {
                Document doc = wordApp.Documents.Open(file);
                try
                {
                    RemoveBibliography(doc);

                  /*  string[] referenceTypes = {   
                        "*. [(][0-9]{4}[)].*",
                        "*. [(][0-9]{4}[!)]@[)].*",
                        "*. [(]n.d.[)].*"
                    }; */
                  

                    // Combine all patterns into a single array
                    string[][] allPatterns = { referenceTypes, inTextReferenceTypes, replaceSymbols, removeSymbols };

                    Parallel.ForEach(doc.StoryRanges.Cast<Range>(), range =>
                    {
                        foreach (string[] patterns in allPatterns)
                        {
                            bool matchWildcards = patterns == referenceTypes || patterns == inTextReferenceTypes;
                            Find findObject = range.Find;
                            findObject.ClearFormatting();
                            findObject.Replacement.ClearFormatting();
                            findObject.Replacement.Text = " ";

                            foreach (string pattern in patterns)
                            {
                                findObject.Text = pattern;
                                findObject.MatchWildcards = matchWildcards;
                                findObject.Execute(Replace: WdReplace.wdReplaceAll);
                            }
                        }
                    }); 

                    doc.SaveAs(file + "modified.docx");
                    Console.WriteLine("Processed doc: " + file);
                }
                finally
                {
                    doc.Close(false);
                }
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
                rng.Start = rng.End;
                rng.End = doc.Content.End;
                rng.Delete();
            }
        }
    }
}
