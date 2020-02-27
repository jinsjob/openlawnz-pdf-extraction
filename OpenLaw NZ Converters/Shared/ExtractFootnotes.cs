using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Shared
{

    class PageFootnoteData
    {
        public int PageNumber { get; set; }
        public string FootnoteText { get; set; }
    }
    public class ExtractFootnotes
    {

        public static string Timestamp
        {
            get
            {
                return new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds().ToString();
            }
        }
        public static object[] jsonOutput = new object[3];
        static string footNoteContexts = string.Empty;
        public static void Process(Document Doc, Application application, Logger logger)
        {
            Document activeDocument = Doc;

            Range makeSearchRange()
            {
                Range searchRange = application.ActiveDocument.Range();

                //searchRange.Find.MatchWholeWord = true;
                searchRange.Find.Format = true;
                return searchRange;
            }

            List<string> conjoinedFootnotes = new List<string>();

            int footnoteStartNumber = 1;

            string footnotesPath = activeDocument.FullName.Replace(".docx", ".footnotes.txt");
            string footnoteContextsPath = activeDocument.FullName.Replace(".docx", ".footnotecontexts.txt").Replace(".pdf", ".footnotecontexts.txt");

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
                //using (StreamWriter sw = File.AppendText(footnotesPath))
                //{
                //sw.WriteLine(data.Trim());
                if (string.IsNullOrEmpty(data))
                    jsonOutput[0] = false;
                else
                    jsonOutput[0] = true;
                jsonOutput[1] = data;

                //  }
            }

            void logFootnoteContextData(string data)
            {
                //using (StreamWriter sw = File.AppendText(footnoteContextsPath))
                //{
                //sw.WriteLine(data.Trim());
                footNoteContexts = footNoteContexts + "," + data;

                // }
            }


            // ------------------------------------------------------------------------------------------
            //
            // If document is in __openlawnz_from_pdf folder or has it in its name
            //
            // ------------------------------------------------------------------------------------------


            List<PageFootnoteData> footnotesData = new List<PageFootnoteData>();


            // First find all the numbers for footnotes so that we know how many there are
            int currentAllNumber = 1;
            int allNumberThreshold = 5; // How far out of footnote sync we can go to check if there's more
            int currentAllNumberThreshold = 0;
            Boolean hasFootnoteIssue = false;

            Range allFootnoteNumbers = makeSearchRange();
            allFootnoteNumbers.Find.Text = currentAllNumber + "";
            allFootnoteNumbers.Find.Font.Size = (float)6.5;
            allFootnoteNumbers.Find.MatchWholeWord = true;
            allFootnoteNumbers.Find.Execute();

            logger.log("Start footnote number check");

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
                allFootnoteNumbers.Collapse(WdCollapseDirection.wdCollapseEnd);
                allFootnoteNumbers.Find.Execute();
            }

            logger.log("End footnote number check");

            logger.log("hasFootnoteIssue: " + hasFootnoteIssue);

            if (hasFootnoteIssue)
            {
                return;
            }

            Range footnoteStartRange = makeSearchRange();

            int currentNumber = 1;

            footnoteStartRange.Find.Text = currentNumber + "";
            footnoteStartRange.Find.Font.Size = (float)6.5;
            //footnoteStartRange.Find.MatchPrefix = true;
            footnoteStartRange.Find.MatchWholeWord = true;
            footnoteStartRange.Find.Execute();

            logger.log("Start footnote search");

            while (footnoteStartRange.Find.Found)
            {

                int currentPage = footnoteStartRange.Information[WdInformation.wdActiveEndPageNumber];

                Range footnoteRange;

                if (currentPage != activeDocument.Content.Information[WdInformation.wdNumberOfPagesInDocument])
                {
                    footnoteRange = application.ActiveDocument.Range(footnoteStartRange.Start, footnoteStartRange.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToNext, 1).Start);
                }
                else
                {
                    footnoteRange = application.ActiveDocument.Range(footnoteStartRange.Start, activeDocument.Content.Characters.Last.End);
                }
                footnoteRange.Font.Shading.BackgroundPatternColor = WdColor.wdColorBlue;

                footnotesData.Add(new PageFootnoteData
                {
                    PageNumber = currentPage,
                    FootnoteText = footnoteRange.Text
                });


                Range searchRange = application.ActiveDocument.Range(footnoteRange.Start, footnoteRange.End);

                searchRange.Find.Text = currentNumber + "";
                searchRange.Find.Font.Size = (float)6.5;
                searchRange.Find.Format = true;

                searchRange.Find.Execute();

                bool findLastFootnoteNumber = true;

                while (searchRange.Find.Found && findLastFootnoteNumber)
                {

                    if (searchRange.Information[WdInformation.wdActiveEndPageNumber] == currentPage)
                    {
                        currentNumber++;
                        searchRange.Find.Text = currentNumber + "";
                        searchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                        searchRange.Find.Execute();
                    }
                    else
                    {
                        findLastFootnoteNumber = false;
                    }
                }

                footnoteStartRange.Find.Text = currentNumber + "";
                footnoteStartRange.Collapse(WdCollapseDirection.wdCollapseEnd);

                footnoteStartRange.Find.Execute();

            }

            logger.log("End footnote search");

            logger.log("footnotesData.Count: " + footnotesData.Count);

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
                int footnoteCurrentNumber = footnoteStartNumber;
                int currentIndex = 0;
                PageFootnoteData currentFootnote = null;

                allFootnotes.ForEach(footnoteData =>
                {
                    // If we are midway through a footnote, append to it
                    if (
                        currentFootnote != null &&
                        !footnoteData.FootnoteText.StartsWith((footnoteCurrentNumber + 1).ToString()) &&    // Does not start with the next footnote number. Can't check for just number, since line could start with a date
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
                        footnoteCurrentNumber++;
                    }

                    if (footnoteData.FootnoteText.StartsWith(footnoteCurrentNumber.ToString()))
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
                return;
            }

            // Find contexts
            // They should be added in order

            Dictionary<int, string> footnoteContextsData = new Dictionary<int, string>();

            // Helper function to generate contexts
            void addContext(Range range)
            {
                Range contextRange = activeDocument.Range(Start: range.Start - 20, End: range.Start + range.Text.Length);
                _ = int.TryParse(range.Text, out int n);

                contextRange.Select();
                // https://stackoverflow.com/questions/238002/replace-line-breaks-in-a-string-c-sharp
                footnoteContextsData.Add(n, contextRange.Text.Trim().Replace("\r\n", "").Replace("\n", "").Replace("\r", ""));
            }

            int currentContextNumber = 1;
            int maximumNumber = conjoinedFootnotes.Count;

            logger.log("Start footnote context search");
            logger.log("maximumNumber: " + maximumNumber);

            while (maximumNumber >= currentContextNumber)
            {
                logger.log(currentContextNumber + "/" + maximumNumber);
                // Find superscript

                Range superscriptSearchRange = makeSearchRange();

                bool superscriptFound = false;
                bool fontSizeFound = false;

                superscriptSearchRange.Find.Font.Superscript = 1;
                superscriptSearchRange.Find.Text = currentContextNumber + "";

                superscriptSearchRange.Find.Execute();

                logger.log("superscript search start");

                if (superscriptSearchRange.Find.Found)
                {
                    addContext(superscriptSearchRange);
                    superscriptFound = true;
                }

                logger.log("superscript search end");
                logger.log("superscriptFound: " + superscriptFound);
                // Check if on same page
                // Log it. Sorted with multiple
                if (!superscriptFound)
                {

                    // Find within font range

                    int fontSizeMin = 5;
                    int fontSizeMax = 8;
                    double currentFontSize = fontSizeMin;

                    logger.log("fontSize search start");

                    while (fontSizeMax >= currentFontSize)
                    {

                        logger.log(currentFontSize + "/" + fontSizeMax);

                        Range fontSizeSearchRange = makeSearchRange();

                        fontSizeSearchRange.Find.Text = currentContextNumber + "";
                        fontSizeSearchRange.Find.Font.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic;
                        fontSizeSearchRange.Find.Font.Size = (float)currentFontSize;


                        fontSizeSearchRange.Find.Execute();

                        logger.log("start fontSizeSearchRange.Find.Found: " + fontSizeSearchRange.Find.Found);

                        while (fontSizeSearchRange.Find.Found)
                        {
                            logger.log("fontSizeSearchRange.Find.Found: " + fontSizeSearchRange.Find.Found);

                            if (fontSizeSearchRange.Font.Position > 2)
                            {
                                addContext(fontSizeSearchRange);
                                fontSizeFound = true;

                                logger.log("fontSizeFound: " + fontSizeFound);
                                break;
                            }
                            fontSizeSearchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                            fontSizeSearchRange.Find.Execute();
                        }

                        logger.log("end fontSizeSearchRange.Find.Found. fontSizeFound: " + fontSizeFound);

                        if (fontSizeFound)
                        {
                            break;
                        }

                        currentFontSize += 0.5;
                    }

                }

                logger.log("end number");
                logger.log("superscriptFound: " + superscriptFound);
                logger.log("fontSizeFound: " + fontSizeFound);

                if (!superscriptFound && !fontSizeFound)
                {
                    break;
                }

                currentContextNumber++;

            }

            logger.log("End footnote context search");

            logger.log("footnoteContextsData.Count: " + footnoteContextsData.Count);

            if (footnoteContextsData.Count > 0)
            {

                foreach (KeyValuePair<int, string> footnoteContext in footnoteContextsData)
                {
                    logFootnoteContextData(footnoteContext.Value);
                }
                jsonOutput[2] = footNoteContexts.Remove(0, 1);
            }

            logger.log("finished. SaveAs2");
            logger.log("closing");
            logger.log("quitting");

        }


    }
}
