using System;
using System.IO;

namespace XLSConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputDir, outputDir;

                if (args.Length != 2)
                {
                    Console.WriteLine("This program batch converts xls and xlsx files to xlsb.\n" +
                        "Folder traversal is recursive.\nCmd Line Usage: Binarizer.exe <inputDirectory> <outputDirectory>\n");
                    Console.WriteLine("Please enter input directory:");
                    inputDir = Console.ReadLine().Replace("\"", "");

                    Console.WriteLine("Output directory:");
                    outputDir = Console.ReadLine().Replace("\"", "");
                }
                else
                {
                    inputDir = args[0];
                    outputDir = args[1];
                }

                inputDir = PathOperations.VerifyRootedPath(inputDir);
                outputDir = PathOperations.VerifyRootedPath(outputDir);

                Console.WriteLine("Writing to " + outputDir);

                FileInfo[] toConvert = CollectFilesForConversion(inputDir);
                PrepareOutputDirectory(outputDir);

                using (ExcelInterop excel = new ExcelInterop())
                {
                    excel.ConvertFiles(toConvert, inputDir, outputDir);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + " Exiting program...press enter to close");
                Console.ReadLine();
            }
        }

        private static FileInfo[] CollectFilesForConversion(string inputDir)
        {
            DirectoryInfo inDir = new DirectoryInfo(inputDir);
            if (inDir.Exists)
            {
                return inDir.GetFiles("*.xls", SearchOption.AllDirectories);
            }
            else
            {
                throw new Exception(string.Format("Input directory ({0}) not found.", inputDir));
            }
        }

        private static void PrepareOutputDirectory(string outputDir)
        {
            DirectoryInfo dir = new DirectoryInfo(outputDir);
            if (dir.Exists)
            {
                if (dir.GetFiles("*", SearchOption.AllDirectories).Length > 0)
                {
                    Console.WriteLine("Output directory is not empty. Overwrites may occur. Proceed? (Y/N)");
                    string overwrite = Console.ReadLine();
                    if (!overwrite.ToUpper().Equals("Y"))
                    {
                        throw new Exception("");
                    }
                }
            }
            else
            {
                dir.Create();
            }

            Logging.Configure(outputDir);
        }
    }
}
