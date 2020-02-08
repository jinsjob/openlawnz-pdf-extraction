using System;
using System.IO;

namespace Shared
{
    public class Logger
    {
        private string filePath;

        public Logger(string filePath)
        {
            this.filePath = filePath;
        }

        public void log(string message, Boolean alsoWriteToConsole = false)
        {
            message = DateTime.Now + "\t" + message.Trim();
            if (alsoWriteToConsole)
            {
                Console.WriteLine(message);
            }
            using (StreamWriter sw = File.AppendText(this.filePath))
            {
                sw.WriteLine(message);
            }
        }
    }
}
