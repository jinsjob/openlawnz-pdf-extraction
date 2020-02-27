using McMaster.Extensions.CommandLineUtils;
using Microsoft.Office.Interop.Word;
using Shared;
using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Text.Json;

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
        public string case_text { get; set; }
        public bool footnotes_present { get; set; }
        public string footnotes { get; set; }
        public string footnote_contexts { get; set; }
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
                objJsonOutput.case_text = document.Content.Text;
                document.Close(false);
                Marshal.ReleaseComObject(document);
                document = application.Documents.Open(filePath);
                ExtractFootnotes.Process(document, application, documentLogger);
                objJsonOutput.footnotes_present = Convert.ToBoolean(ExtractFootnotes.jsonOutput[0].ToString());
                objJsonOutput.footnotes = ExtractFootnotes.jsonOutput[1].ToString();

                objJsonOutput.footnote_contexts = ExtractFootnotes.jsonOutput[2].ToString();
                File.WriteAllText(jsonPath, JsonSerializer.Serialize<JsonOutput>(objJsonOutput));

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
