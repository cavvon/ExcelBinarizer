using System;
using System.Collections.Generic;
using System.IO;

namespace XLSConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("This program batch converts xls and xlsx files to xlsb.\n" +
                    "Folder traversal is recursive.\nCmd Line Usage: Binarizer.exe <inputDirectory> <outputDirectory>\n");
            }

            bool rerun = false;
            do
            {
                rerun = false;
                try
                {
                    string useDefault, inputDir, outputDir;
                    bool rePrompt = true;

                    if (args.Length != 2)
                    {
                        do
                        {
                            Console.WriteLine("Use default directories (.\\input; .\\output) [y/n]?");
                            useDefault = Console.ReadLine();
                            if (useDefault.ToUpper().Equals("Y"))
                            {
                                rePrompt = false;
                                inputDir = "input";
                                outputDir = "output";
                            }
                            else if (useDefault.ToUpper().Equals("N"))
                            {
                                rePrompt = false;
                                Console.WriteLine("Please enter input directory:");
                                inputDir = Console.ReadLine().Replace("\"", "");

                                Console.WriteLine("Output directory:");
                                outputDir = Console.ReadLine().Replace("\"", "");
                            }
                            else
                            {
                                // Remove compiler warnings
                                inputDir = null;
                                outputDir = null;
                            }

                        } while (rePrompt);
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

                Console.WriteLine("Run again? [y/n]");
                string rerunIn = Console.ReadLine();
                if (rerunIn.ToUpper().Equals("Y"))
                {
                    rerun = true;
                }
                else
                {
                    rerun = false;
                }
            } while (rerun);
        }

        private static FileInfo[] CollectFilesForConversion(string inputDir)
        {
            DirectoryInfo inDir = new DirectoryInfo(inputDir);
            List<FileInfo> conversionFiles = new List<FileInfo>();
            if (inDir.Exists)
            {
                conversionFiles.AddRange(inDir.GetFiles("*.xls", SearchOption.AllDirectories)); // Adds xls extensions (xls, xlsx, xlsb...)
                conversionFiles.AddRange(inDir.GetFiles("*.csv", SearchOption.AllDirectories));
                conversionFiles.AddRange(inDir.GetFiles("*.tsv", SearchOption.AllDirectories));
                return conversionFiles.ToArray();
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
