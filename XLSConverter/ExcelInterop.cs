using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace XLSConverter
{
    public class ExcelInterop : IDisposable
    {
        private bool _disposed = false;
        private Application excelApp;

        public ExcelInterop()
        {
            excelApp = new Application();
            excelApp.DisplayAlerts = true;
        }

        public void ConvertFiles(FileInfo[] files, string inputDir, string outputDir)
        {
            int errors = 0;

            for (int i = 0; i < files.Length; i++)
            {
                try
                {
                    string newPath = GenerateNewPath(files[i].FullName, inputDir, outputDir);
                    Convert(files[i].FullName, newPath, excelApp);
                    Console.Write(String.Format("\r{0} / {1} Converted", i + 1, files.Length));
                }
                catch (Exception e)
                {
                    errors++;
                    Logging.Write(files[i].FullName, "ERROR: " + e.Message);
                }
            }

            Console.WriteLine(string.Format("\nDone with {0} errors! Logfile is in output directory.", errors));
            Logging.Write(string.Format("Done with {0} errors. Search ERROR to find.", errors));
            Console.ReadLine();
        }

        private static string GenerateNewPath(string oldPath, string inputDir, string outputDir)
        {
            string subDir = oldPath.Remove(0, inputDir.Length).TrimStart(new char[] { '\\' });
            Logging.Write(subDir);

            string newPath = Path.Combine(outputDir, Path.ChangeExtension(subDir, ".xlsb"));

            // Ensure new output path exists and is available
            string newDir = Path.GetDirectoryName(newPath);

            if (!Directory.Exists(newDir))
            {
                Directory.CreateDirectory(newDir);
            }
            else if (File.Exists(newPath))
            {
                File.Delete(newPath);
            }

            return newPath;
        }

        private static void Convert(string oldPath, string newPath, Application excelApp)
        {
            Workbook workbook = excelApp.Workbooks.Open(oldPath, XlUpdateLinks.xlUpdateLinksNever, true,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // save in XlFileFormat.xlExcel12 format which is XLSB
            workbook.SaveAs(newPath, XlFileFormat.xlExcel12,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            workbook.Close(false, Type.Missing, Type.Missing);
            Marshal.ReleaseComObject(workbook);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {


                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            this._disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
