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
            foreach(string file in files)            
            {
                Document doc = wordApp.Documents.Open(file);
                try
                {
                    RemoveBibliography(doc);



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

                  /*  string[] referenceTypes = {   
                        "*. [(][0-9]{4}[)].*",
                        "*. [(][0-9]{4}[!)]@[)].*",
                        "*. [(]n.d.[)].*"
                    }; */

                    string[] referenceTypes = { @"^13[!^13]@ [(][0-9]{4}[)].[!^13]@^13", @"^13[!^13]@ [(][0-9]{4}[!)]@[)].[!^13]@^13" };
                    

                    // Combine all patterns into a single array
                    string[][] allPatterns = { referenceTypes, inTextReferenceTypes, replaceSymbols, removeSymbols };
                    
                    foreach (Range range in doc.StoryRanges)
                    {
                        foreach (string[] patterns in allPatterns)
                        {
                            foreach (string pattern in patterns)
                            {
                                Find findObject = range.Find;
                                findObject.ClearFormatting();
                                findObject.Text = pattern;
                                findObject.Replacement.ClearFormatting();
                                findObject.Replacement.Text = " ";
                                findObject.MatchWildcards = patterns == referenceTypes || patterns == inTextReferenceTypes;
                                findObject.Execute(Replace: WdReplace.wdReplaceAll);

                            }
                        }
                    }
                    
                    doc.SaveAs(file + "modified.docx");
                    Console.WriteLine("Processed doc");
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
