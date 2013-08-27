using System.IO;
using System;

namespace XLSConverter
{
    public static class Logging
    {
        private static string logPath = "";

        public static void Configure(string outputDir)
        {
            logPath = Path.Combine(outputDir, "log.txt");
            using (FileStream fs = File.Create(logPath)) { }
        }

        public static void Write(string file, string msg)
        {
            Write(string.Format("{0}: {1}", file, msg));
        }

        public static void Write(string msg)
        {
            if (!string.IsNullOrEmpty(logPath))
            {
                using (StreamWriter sw = File.AppendText(logPath))
                {
                    sw.WriteLine(string.Format("{0} - {1}", DateTime.Now.ToString("yyyy-MM-dd h:mm:ss t"), msg));
                }
            }
        }
    }
}
