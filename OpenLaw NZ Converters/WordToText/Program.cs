using McMaster.Extensions.CommandLineUtils;
using Microsoft.Office.Interop.Word;
using Shared;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace WordToText
{
    //TODO: Footnotes present but invalid

    public class DirectorySection : ConfigurationSection
    {
        public static string Name = "Directory";

        [ConfigurationProperty("TargetDirectory", IsRequired = true)]
        public string TargetDirectory
        {
            get
            {
                return (string)this["TargetDirectory"];
            }
            set
            {
                this["TargetDirectory"] = value;
            }
        }

        [ConfigurationProperty("WatchDirectory", IsRequired = false)]
        public Boolean WatchDirectory
        {
            get
            {
                return (Boolean)this["WatchDirectory"];
            }
            set
            {
                this["WatchDirectory"] = value;
            }
        }

    }

    public class JsonOutput
    {
        public List<string> case_text { get; set; }
        public bool footnotes_present { get; set; }
        public string footnotes { get; set; }
        public List<string> footnote_contexts { get; set; }
    }

    class PageFootnoteData
    {
        public int PageNumber { get; set; }
        public string FootnoteText { get; set; }
    }

    class Program
    {
        public static int FilesBeingProcessed = 0;
        static JsonOutput objJsonOutput = new JsonOutput();
        static string jsonPath;

        // TODO: make logger optional
        private static void CleanUpWordInstances(Logger logger)
        {
            var existingWordProcesses = System.Diagnostics.Process.GetProcessesByName("winword");

            if (existingWordProcesses.Length > 0)
            {

                logger.log("Cleanup Word processes", true);

                foreach (var process in System.Diagnostics.Process.GetProcessesByName("winword"))
                {
                    process.Kill();
                }

            }
        }
        public static void ProcessFile(string filePath, Logger parentLogger)
        {

             jsonPath = filePath.Replace(".docx", ".json");
            string logPath = filePath.Replace(".docx", ".log" + ExtractFootnotes.Timestamp + ".txt");

            FilesBeingProcessed++;

            Logger logger = new Logger(logPath);

            logger.log(String.Format("Start Process (Processing count = {0})", FilesBeingProcessed), true);

            Application application = new Application();
            application.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            Logger documentLogger = new Logger(logPath);

            Document document = null;
            try
            {

                document = application.Documents.Open(filePath);

                string text = document.Content.ToString();
                List<string> data = new List<string>();
                foreach (Paragraph objparagraph in document.Paragraphs)
                {
                    data.Add(objparagraph.Range.Text.Trim());
                }

                objJsonOutput.case_text = data;
                document.Close(false);
                Marshal.ReleaseComObject(document);
                document = application.Documents.Open(filePath);
                Process(document, application, documentLogger);

            }
            catch (Exception ex)
            {
                parentLogger.log(ex.ToString() + " " + ex.Message);
            }

            document.Close(false);
            Marshal.ReleaseComObject(document);
            application.Quit(false);
            Marshal.ReleaseComObject(application);

            FilesBeingProcessed--;

            logger.log(String.Format("Finish Process (Processing count = {0})", FilesBeingProcessed), true);

        }
        public static void Run(string targetDirectory, Boolean watchDirectory)
        {

            if (!Directory.Exists(targetDirectory))
            {
                Console.WriteLine("Directory does not exist. Either modify the configuration file, or run this program again and re-save the configuration. Press any key to exit.");
                Console.ReadLine();
                return;
            }

            Logger logger = new Logger(Path.Combine(targetDirectory, String.Format("_log-{0}.txt", ExtractFootnotes.Timestamp)));

            CleanUpWordInstances(logger);

            var previousConversionTextFiles = Directory.EnumerateFiles(targetDirectory, "*.txt").Where(f => !f.Contains("_log"));

            if (previousConversionTextFiles.Count() > 0)
            {
                logger.log("Cleanup conversion text files", true);

                foreach (string textFile in previousConversionTextFiles)
                {
                    File.Delete(textFile);
                }

            }

            var previousConversionWordTemporaryFiles = Directory.EnumerateFiles(targetDirectory, "*.docx").Where(f => f.Contains("~$o_"));

            if (previousConversionWordTemporaryFiles.Count() > 0)
            {
                logger.log("Cleanup conversion Word temporary files", true);

                foreach (string wordFile in previousConversionWordTemporaryFiles)
                {
                    File.Delete(wordFile);
                }

            }

            logger.log("Start", true);

            var files = Directory.EnumerateFiles(targetDirectory, "*.docx").Where(f => !f.Contains("~$o_"));

            logger.log(String.Format("The number of files is {0}.", files.Count()), true);

            Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 3 }, file =>
            {
                ProcessFile(file, logger);
            });

            Marshal.CleanupUnusedObjectsInCurrentContext();

            CleanUpWordInstances(logger);

            logger.log("Finish", true);
        }

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

            void logFootnoteData(string data)
            {
                objJsonOutput.footnotes_present = true;
                objJsonOutput.footnotes = JsonSerializer.Serialize<string>(data);
            }

            List<string> footnoteContexts = new List<string>();
            void logFootnoteContextData(string data)
            {
                footnoteContexts.Add(data);

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
            }
            objJsonOutput.footnote_contexts = footnoteContexts;
            File.WriteAllText(jsonPath, JsonSerializer.Serialize<JsonOutput>(objJsonOutput));

            logger.log("finished. SaveAs2");
            logger.log("closing");
            logger.log("quitting");

        }
        public static int Main(string[] args)
        {

            // If there is a configuration, use that

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            Console.WriteLine(String.Format("Using configuration file at {0}", config.FilePath));

            var directorySection = (DirectorySection)config.GetSection(DirectorySection.Name);

            if (directorySection != null && !string.IsNullOrEmpty(directorySection.TargetDirectory))
            {
                Run(directorySection.TargetDirectory, directorySection.WatchDirectory);
                return 0;

            }
            else
            {
                var app = new CommandLineApplication();

                app.HelpOption();

                app.OnExecute(() =>
                {

                    string newTargetDirectory = null;
                    Boolean newWatchDirectory = false;
                    Boolean newSaveConfiguration = false;

                    Console.WriteLine("Debugger prompts because no config set:");

                    while (string.IsNullOrEmpty(newTargetDirectory))
                    {
                        string potentialNewTargetDirectory = Prompt.GetString("Enter a directory path");

                        if (!string.IsNullOrEmpty(potentialNewTargetDirectory) && Directory.Exists(potentialNewTargetDirectory))
                        {
                            newTargetDirectory = potentialNewTargetDirectory;
                        }
                        else
                        {
                            Console.WriteLine("Directory does not exist");
                        }
                    }
                    newWatchDirectory = Prompt.GetYesNo("Watch directory?", false);
                    newSaveConfiguration = Prompt.GetYesNo("Save configuration?", true);

                    if (newSaveConfiguration)
                    {
                        Console.WriteLine("Saving configuration");

                        config.Sections.Add(DirectorySection.Name, new DirectorySection
                        {
                            TargetDirectory = newTargetDirectory,
                            WatchDirectory = newWatchDirectory
                        });
                        // Save the configuration file.
                        config.Save(ConfigurationSaveMode.Modified);

                        // Force a reload of the changed section. This 
                        // makes the new values available for reading.
                        ConfigurationManager.RefreshSection(DirectorySection.Name);

                    }

                    Run(newTargetDirectory, newSaveConfiguration);

                });
                return app.Execute(args);

            }
        }

    }
}
