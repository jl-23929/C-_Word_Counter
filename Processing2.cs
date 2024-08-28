using System;
using OfficeIMO.Word;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Spreadsheet;
using Spire.Doc;
using System.Collections.Generic;
using System.Linq;

namespace Word_Counter.Processing2
{
    public class Processing2
    {
        public void RemoveSymbols(string[] files)
        {
            string[] patterns = {
        "1", "2", "3", "4", "5", "6", "7", "8", "9", "0",
        ",", ".", "?", "!", ":", ";", "(", ")", "[", "]", "{", "}",
        "/", @"\", "*", "+", "=", "|", "&", "^", "%", "@", "~",
        "`", "'", "°", "𝜃", "×", "±", "≈", "∆", ">", "<", ">=",
        "<=", "=", "ϕ", "φ", "Φ", "Ω", "Ω", "∑", "∞", "√"
    };
            string replacement = " ";

            // Escape each pattern
            string[] escapedPatterns = patterns.Select(Regex.Escape).ToArray();
            string combinedPattern = string.Join("|", escapedPatterns);
            Regex regex = new Regex(combinedPattern);

            foreach (string file in files)
            {
                using (var document = WordDocument.Load(file))
                {
                    // Create a single regex pattern to match all symbols
                    // Iterate through paragraphs and perform replacement
                    foreach (var paragraph in document.Paragraphs)
                    {
                        if (!string.IsNullOrEmpty(paragraph.Text))
                        {
                            paragraph.Text = regex.Replace(paragraph.Text, replacement);
                        }
                    }

                    bool foundBibliography = false;
                    List<int> paragraphsToRemove = new List<int>();

                    for (int i = 0; i < document.Paragraphs.Count; i++)
                    {
                        var paragraph = document.Paragraphs[i];

                        if (foundBibliography)
                        {
                            paragraphsToRemove.Add(i);
                        }
                        else if (!string.IsNullOrEmpty(paragraph.Text) && paragraph.Text.Contains("Bibliography"))
                        {
                            foundBibliography = true;
                        }
                    }

                    // Remove paragraphs in a single pass
                    for (int i = paragraphsToRemove.Count - 1; i >= 0; i--)
                    {
                        document.Paragraphs.RemoveAt(paragraphsToRemove[i]);
                    }

                    // Save the document
                    document.Save(file + "_modified.docx");
                }
            }
        }
    }
}