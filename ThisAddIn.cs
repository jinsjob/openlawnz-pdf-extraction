using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace CaseDataExtractor
{

    public class PageFootnoteData
    {
        public int PageNumber { get; set; }
        public string FootnoteText { get; set; }
    }

    public partial class ThisAddIn
    {

        Word.Range makeSearchRange()
        {
            Word.Range searchRange = this.Application.ActiveDocument.Range();

            //searchRange.Find.MatchWholeWord = true;
            searchRange.Find.Format = true;
            return searchRange;
        }

        private void processDocument(Word.Document Doc)
        {
            Word.Document activeDocument = Doc;

            List<string> conjoinedFootnotes = new List<string>();

            int footnoteStartNumber = 1;

            string footnotesPath = activeDocument.FullName.Replace(".docx", ".footnotes.txt").Replace(".pdf", ".footnotes.txt");
            string footnoteContextsPath = activeDocument.FullName.Replace(".docx", ".footnotecontexts.txt").Replace(".pdf", ".footnotecontexts.txt");
            string fullTextPath = activeDocument.FullName.Replace(".docx", ".txt");

            string[] splitPaths = footnotesPath.Split(new string[] { "split-" }, StringSplitOptions.None);

            if (splitPaths.Length > 1)
            {

                char splitPath = splitPaths.Last().First();

                int existingNumber = Int32.Parse(splitPath.ToString());

                string newPath = Regex.Replace(footnotesPath, @"split-\d+", "split-" + (existingNumber - 1));
                // Get previous footnote
                if (File.Exists(newPath))
                {
                    // c# read last line and first numbers
                    string[] previousFootnotes = File.ReadAllLines(newPath);

                    footnoteStartNumber = Int32.Parse(Regex.Match(previousFootnotes.Last(), @"^\d+").Value);

                }

            };



            if (File.Exists(footnotesPath))
            {
                File.Delete(footnotesPath);
            }

            if (File.Exists(footnoteContextsPath))
            {
                File.Delete(footnoteContextsPath);
            }

            void logFootnoteData(string data)
            {
                using (StreamWriter sw = File.AppendText(footnotesPath))
                {
                    sw.WriteLine(data.Trim());
                }
            }

            void logFootnoteContextData(string data)
            {
                using (StreamWriter sw = File.AppendText(footnoteContextsPath))
                {
                    sw.WriteLine(data.Trim());
                }
            }


            // ------------------------------------------------------------------------------------------
            //
            // If document is in __openlawnz_from_pdf folder or has it in its name
            //
            // ------------------------------------------------------------------------------------------
            if (activeDocument.FullName.Contains("__openlawnz_from_pdf"))
            {

                Application.ScreenUpdating = false;
                Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                List<PageFootnoteData> footnotesData = new List<PageFootnoteData>();

                // Slow
                if (activeDocument.FullName.Contains("__use_shapes"))
                {
                    // Find line shapes on the page because chances are footnotes are under them
                    // Take the text under the shape and append to a string

                    SortedDictionary<int, Word.Shape> pageShapesDictionary = new SortedDictionary<int, Word.Shape>();
                    foreach (Word.Shape pageShape in activeDocument.Shapes)
                    {
                        int sPageNumber = pageShape.Anchor.Information[Word.WdInformation.wdActiveEndPageNumber];
                        if (pageShapesDictionary.ContainsKey(sPageNumber))
                        {
                            // If multiple images, get the one furthest down the page

                            if (pageShape.Anchor.Information[Word.WdInformation.wdVerticalPositionRelativeToPage] > pageShapesDictionary[sPageNumber].Anchor.Information[Word.WdInformation.wdVerticalPositionRelativeToPage])
                            {
                                pageShapesDictionary.Remove(sPageNumber);
                                pageShapesDictionary.Add(sPageNumber, pageShape);
                            }
                        }
                        else
                        {
                            pageShapesDictionary.Add(sPageNumber, pageShape);
                        }
                    }


                    SortedDictionary<int, double> footnoteShapePositions = new SortedDictionary<int, double>();

                    // Assuming one line per page after the first page

                    foreach (KeyValuePair<int, Word.Shape> pageShape in pageShapesDictionary)
                    {

                        double sTopPosition = pageShape.Value.Anchor.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];

                        // From the first shape that is the correct dimension, add
                        if (footnoteShapePositions.Count > 0 || (pageShape.Value.Width < 4084 && pageShape.Value.Left < 200))
                        {
                            footnoteShapePositions.Add(pageShape.Key, sTopPosition);
                        }
                    }

                    //this.Application.Quit();


                    // Go through all paragraphs and if they are under a footnote shape, add

                    foreach (Word.Paragraph p in activeDocument.Paragraphs)
                    {
                        double r = p.Range.Information[Word.WdInformation.wdVerticalPositionRelativeToPage];

                        foreach (KeyValuePair<int, double> footnoteShape in footnoteShapePositions)
                        {
                            if (footnoteShape.Key == p.Range.Information[Word.WdInformation.wdActiveEndPageNumber])
                            {
                                if (r > footnoteShape.Value)
                                {

                                    if (int.TryParse(p.Range.Characters.First.Text, out int n))
                                    {
                                        // Possibly footnote
                                        p.Range.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlue;
                                        // Footnotes are added in a bulk block of text, not line by line

                                        footnotesData.Add(new PageFootnoteData
                                        {
                                            PageNumber = footnoteShape.Key,
                                            FootnoteText = p.Range.Text
                                        });

                                    }
                                }
                            }
                        }
                    }

                }
                else
                {

                    // First find all the numbers for footnotes so that we know how many there are
                    int currentAllNumber = 1;
                    int allNumberThreshold = 5; // How far out of footnote sync we can go to check if there's more
                    int currentAllNumberThreshold = 0;
                    Boolean hasFootnoteIssue = false;

                    Word.Range allFootnoteNumbers = makeSearchRange();
                    allFootnoteNumbers.Find.Text = currentAllNumber + "";
                    allFootnoteNumbers.Find.Font.Size = (float)6.5;
                    allFootnoteNumbers.Find.MatchWholeWord = true;
                    allFootnoteNumbers.Find.Execute();

                    while ((allFootnoteNumbers.Find.Found || currentAllNumberThreshold < allNumberThreshold) && !hasFootnoteIssue)
                    {
                        if (allFootnoteNumbers.Find.Found)
                        {
                            hasFootnoteIssue = currentAllNumberThreshold > 0;
                        }
                        else
                        {
                            currentAllNumberThreshold++;
                        }
                        currentAllNumber++;
                        allFootnoteNumbers.Find.Text = currentAllNumber + "";
                        allFootnoteNumbers.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        allFootnoteNumbers.Find.Execute();
                    }

                    if (hasFootnoteIssue)
                    {
                        activeDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                        Application.Quit();
                        return;
                    }

                    Word.Range footnoteStartRange = makeSearchRange();

                    int currentNumber = 1;

                    footnoteStartRange.Find.Text = currentNumber + "";
                    footnoteStartRange.Find.Font.Size = (float)6.5;
                    //footnoteStartRange.Find.MatchPrefix = true;
                    footnoteStartRange.Find.MatchWholeWord = true;
                    footnoteStartRange.Find.Execute();


                    while (footnoteStartRange.Find.Found)
                    {

                        int currentPage = footnoteStartRange.Information[Word.WdInformation.wdActiveEndPageNumber];

                        Word.Range footnoteRange;

                        if (currentPage != activeDocument.Content.Information[Word.WdInformation.wdNumberOfPagesInDocument])
                        {
                            footnoteRange = this.Application.ActiveDocument.Range(footnoteStartRange.Start, footnoteStartRange.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToNext, 1).Start);
                        }
                        else
                        {
                            footnoteRange = this.Application.ActiveDocument.Range(footnoteStartRange.Start, activeDocument.Content.Characters.Last.End);
                        }
                        footnoteRange.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorBlue;

                        footnotesData.Add(new PageFootnoteData
                        {
                            PageNumber = currentPage,
                            FootnoteText = footnoteRange.Text
                        });


                        Word.Range searchRange = this.Application.ActiveDocument.Range(footnoteRange.Start, footnoteRange.End);

                        searchRange.Find.Text = currentNumber + "";
                        searchRange.Find.Font.Size = (float)6.5;
                        searchRange.Find.Format = true;

                        searchRange.Find.Execute();

                        bool findLastFootnoteNumber = true;

                        while (searchRange.Find.Found && findLastFootnoteNumber)
                        {

                            if (searchRange.Information[Word.WdInformation.wdActiveEndPageNumber] == currentPage)
                            {
                                currentNumber++;
                                searchRange.Find.Text = currentNumber + "";
                                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                searchRange.Find.Execute();
                            }
                            else
                            {
                                findLastFootnoteNumber = false;
                            }
                        }

                        footnoteStartRange.Find.Text = currentNumber + "";
                        footnoteStartRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        footnoteStartRange.Find.Execute();

                    }

                }

                if (footnotesData.Count > 0)
                {

                    //footnotesData.Sort((PosA, PosB) => PosA.PageNumber.CompareTo(PosB.PageNumber));

                    List<PageFootnoteData> allFootnotes = new List<PageFootnoteData>();
                    var newLinesRegex = new Regex(@"\r\n|[\n\v\f\r\x85\u2028\u2029]", RegexOptions.Singleline);

                    footnotesData.ForEach(f =>
                    {
                        string[] splitFootnoteText = newLinesRegex.Split(f.FootnoteText.Trim());

                        foreach (string splitFootnote in splitFootnoteText)
                        {
                            allFootnotes.Add(new PageFootnoteData
                            {
                                PageNumber = f.PageNumber,
                                FootnoteText = splitFootnote
                            });
                        }

                    });

                    int footnotesArrayLen = allFootnotes.Count;
                    int currentNumber = footnoteStartNumber;
                    int currentIndex = 0;
                    PageFootnoteData currentFootnote = null;

                    allFootnotes.ForEach(footnoteData =>
                    {
                        // If we are midway through a footnote, append to it
                        if (
                                currentFootnote != null &&
                                !footnoteData.FootnoteText.StartsWith((currentNumber + 1).ToString()) &&    // Does not start with the next footnote number. Can't check for just number, since line could start with a date
                                footnoteData.PageNumber == currentFootnote.PageNumber                       // Is on the same page
                            )
                        {
                            currentFootnote.FootnoteText += ' ' + footnoteData.FootnoteText;
                        }
                        // If we have a current footnote
                        else if (currentFootnote != null)
                        {
                            conjoinedFootnotes.Add(currentFootnote.FootnoteText);
                            currentFootnote = null;
                            currentNumber++;
                        }

                        if (footnoteData.FootnoteText.StartsWith(currentNumber.ToString()))
                        {
                            currentFootnote = footnoteData;
                        }
                        if (currentFootnote != null && currentIndex == footnotesArrayLen - 1)
                        {
                            conjoinedFootnotes.Add(currentFootnote.FootnoteText);
                        }

                        currentIndex++;
                    });


                    logFootnoteData(string.Join("\n", conjoinedFootnotes));
                }
                else
                {
                    activeDocument.SaveAs2(fullTextPath, Word.WdSaveFormat.wdFormatUnicodeText);
                    activeDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Application.Quit();
                    return;
                }

                // Find contexts
                // They should be added in order

                Dictionary<int, string> footnoteContextsData = new Dictionary<int, string>();

                // Helper function to generate contexts
                void addContext(Word.Range range)
                {
                    Word.Range contextRange = activeDocument.Range(Start: range.Start - 20, End: range.Start + range.Text.Length);
                    _ = int.TryParse(range.Text, out int n);

                    contextRange.Select();
                    // https://stackoverflow.com/questions/238002/replace-line-breaks-in-a-string-c-sharp
                    footnoteContextsData.Add(n, contextRange.Text.Trim().Replace("\r\n", "").Replace("\n", "").Replace("\r", ""));
                }



                int currentContextNumber = 1;
                int maximumNumber = conjoinedFootnotes.Count;

                while (maximumNumber >= currentContextNumber)
                {

                    // Find superscript

                    Word.Range superscriptSearchRange = makeSearchRange();

                    bool superscriptFound = false;
                    bool fontSizeFound = false;

                    superscriptSearchRange.Find.Font.Superscript = 1;
                    superscriptSearchRange.Find.Text = currentContextNumber + "";

                    superscriptSearchRange.Find.Execute();

                    if (superscriptSearchRange.Find.Found)
                    {
                        addContext(superscriptSearchRange);
                        superscriptFound = true;
                    }
                    // Check if on same page
                    // Log it. Sorted with multiple
                    if (!superscriptFound)
                    {

                        // Find within font range

                        int fontSizeMin = 5;
                        int fontSizeMax = 8;
                        double currentFontSize = fontSizeMin;


                        while (fontSizeMax >= currentFontSize)
                        {

                            Word.Range fontSizeSearchRange = makeSearchRange();

                            fontSizeSearchRange.Find.Text = currentContextNumber + "";
                            fontSizeSearchRange.Find.Font.Shading.BackgroundPatternColor = Word.WdColor.wdColorAutomatic;
                            fontSizeSearchRange.Find.Font.Size = (float)currentFontSize;


                            fontSizeSearchRange.Find.Execute();

                            while (fontSizeSearchRange.Find.Found)
                            {

                                if (fontSizeSearchRange.Font.Position > 2)
                                {
                                    addContext(fontSizeSearchRange);
                                    fontSizeFound = true;
                                    break;
                                }
                                fontSizeSearchRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                fontSizeSearchRange.Find.Execute();
                            }

                            if (fontSizeFound)
                            {
                                break;
                            }

                            currentFontSize += 0.5;
                        }

                    }

                    if (!superscriptFound && !fontSizeFound)
                    {
                        break;
                    }

                    currentContextNumber++;

                }
                if (footnoteContextsData.Count > 0)
                {

                    foreach (KeyValuePair<int, string> footnoteContext in footnoteContextsData)
                    {
                        logFootnoteContextData(footnoteContext.Value);
                    }
                }

                activeDocument.SaveAs2(fullTextPath, Word.WdSaveFormat.wdFormatUnicodeText);

                activeDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Application.Quit();
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Word.ApplicationEvents2_Event wdEvents2 = (Word.ApplicationEvents2_Event)this.Application;
            wdEvents2.DocumentOpen += processDocument;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
