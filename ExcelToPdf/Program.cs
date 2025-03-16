using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.WebSockets;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;


//NuGet
using Microsoft.Office.Interop.Excel;

namespace ExcelToPdf
{
    internal class Program
    { 
        static void Main(string[] args)
        {
            ////Parameters////
            string directoryPathFiles = GetPathToDirectory("Files");
            string baseDirectoryPath = directoryPathFiles + @"\BaseFiles";
            string newDirectoryPath = directoryPathFiles + @"\NewFiles";
 
            CreateNewDirectory(newDirectoryPath);

            string[] excelList = ReturnPathForFilesWithExtensionFromDirectory(baseDirectoryPath, ".xlsx");
            string[] pdfList = PreparePathForFilesInNewDirectory(excelList, newDirectoryPath, ".pdf");

            ExcelToPdf(excelList, pdfList);

            Console.WriteLine("\n"+"All finished, to close app press any key");
            Console.ReadKey();

            return;
        }
        static string[] ReturnPathForFilesWithExtensionFromDirectory(string directoryPath, string fileExtension)
        {
            Queue<string> excelFilesQueue = new Queue<string>();
            string[] excelFilesList = null;


            if ('.' != fileExtension[0])
            {
                fileExtension = '.' + fileExtension;
            }

            try
            {
                Console.WriteLine("\ntest");
                Console.WriteLine(directoryPath);
                Console.WriteLine(fileExtension);
                Console.WriteLine("test\n");

                excelFilesList = Directory.GetFiles(directoryPath, fileExtension);
                if (0 == excelFilesList.Length)
                {
                    Console.WriteLine($"\nNo files with extension: '{fileExtension}' in directory:\n" + directoryPath);
                    CloseApp();
                }
            }
            catch
            {
                Debug.WriteLine("ReadPathOfBaseFiles => Directory.GetFiles");

                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }

            bool temporaryFileExist = false;
            for (int i = 0; i < excelFilesList.Length; i++)
            {
                if ('~' != excelFilesList[i].Substring(excelFilesList[i].LastIndexOf(@"\")+1 )[0])
                {
                    Console.WriteLine(excelFilesList[i]);
                    excelFilesQueue.Enqueue(excelFilesList[i]);
                }
                else
                {
                    temporaryFileExist = true;
                }
            }

            if (temporaryFileExist)
            {
                return excelFilesQueue.ToArray();
            }
            else
            {
                return excelFilesList;
            }
        }
        static string[] PreparePathForFilesInNewDirectory(string[] excelList, string newDirectoryPath, string newFileExtension)
        {

            string[] pdfList = new string[excelList.Length];

            try
            {
                string fileExtension = excelList[0].Substring(excelList[0].LastIndexOf(@".") + 1);
                //string fileExtension = excelList[0].Substring(excelList[0].LastIndexOf(@".")+1);
                Console.WriteLine(fileExtension);
                for (int i = 0; i < excelList.Length; i++)
                {
                    pdfList[i] = newDirectoryPath +  
                    excelList[i]
                    .Substring(excelList[i].LastIndexOf(@"\"))
                    .Replace(fileExtension, newFileExtension);
                    
                }
            }
            catch
            {

                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }

            return pdfList;
        }
        static void CreateNewDirectory(string newDirectoryPath)
        {
            try
            {
                if (Directory.Exists(newDirectoryPath))
                {
                    Directory.Delete(newDirectoryPath, true);
                    Directory.CreateDirectory(newDirectoryPath);
                }
                else
                {
                    Directory.CreateDirectory(newDirectoryPath);
                }
            }
            catch (UnauthorizedAccessException)
            {
                Debug.WriteLine("Catch in CreateNewDirectory -> UnauthorizedAccessException");

                Console.WriteLine($"Close all files from directory or check permissions");
                Console.WriteLine("Can't delete directory:\n" + newDirectoryPath);
                CloseApp();
            }
            catch
            {
                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }
        }
        static void ExcelToPdf(string[] excelList, string[] pdfList)
        {
            Application excelApp = null;
            Workbook baseWorkbook = null, newWorkbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Application();

                Console.WriteLine("\nConversion of files from .xlsx to .pdf has started");
                for (int i = 0; i < excelList.Length; i++)
                {
                    baseWorkbook = excelApp.Workbooks.Open(excelList[i]);
                    worksheet = baseWorkbook.Sheets[2];

                    newWorkbook = excelApp.Workbooks.Add();
                    worksheet.Copy(newWorkbook.Sheets[1]);
                    newWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfList[i]);


                    baseWorkbook.Close(false);
                    Marshal.ReleaseComObject(baseWorkbook);
                    baseWorkbook = null;

                    newWorkbook.Close(false);
                    Marshal.ReleaseComObject(newWorkbook);
                    newWorkbook = null;
                    
                    Marshal.ReleaseComObject(worksheet);
                    worksheet = null;

                    Console.WriteLine("File {0} saved: {1}", (i + 1), pdfList[i].Substring(pdfList[i].LastIndexOf(@"\")+1) );
                }

                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;

                Console.WriteLine("All files converted, saved in directory:");
                Console.WriteLine(pdfList[0].Substring(0, pdfList[0].LastIndexOf(@"\") ));
            }
            catch
            {
                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }
            finally
            {
                if (null != excelApp) 
                { 
                    Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                }
                if (null != baseWorkbook)
                {
                    Marshal.ReleaseComObject(baseWorkbook);
                    baseWorkbook = null;
                }
                if (null != newWorkbook) 
                { 
                    Marshal.ReleaseComObject(newWorkbook);
                    newWorkbook = null;
                }
                if (null != worksheet) 
                { 
                    Marshal.ReleaseComObject(worksheet);
                    worksheet = null;
                }
            }

            return;
        }
        static void CloseApp([CallerMemberName] string methodName = "")
        {
            Debug.WriteLine($"\n!!!There was a problem in method: {methodName}\n");

            Console.WriteLine("\nPress Enter to close app");
            Console.ReadKey();

            Environment.Exit(0);
        }
        static string GetPathToDirectory()
        {
            string path = null;
            try
            {
                path = new DirectoryInfo(".").FullName;
            }
            catch
            {
                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }

            int ile = -1;
            try
            {
                ile = path.IndexOf("bin") - 1;
                if (ile < 0)
                {
                    Console.WriteLine("\nPlease contact the IT department");
                    CloseApp();
                }
            }
            catch
            {
                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }


            return path.Substring(0, ile);
        }
        static string GetPathToDirectory(string directoryName)
        {
            string path = null;
            try
            {
                path = new DirectoryInfo(".").FullName;
            }
            catch
            {
                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }

            int ile = -1;
            try
            {
                ile = path.IndexOf("bin") - 1;
                if (ile < 0)
                {
                    Console.WriteLine("\nPlease contact the IT department");
                    CloseApp();
                }
            }
            catch
            {
                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }

            
            return (path.Substring(0, ile) + @"\" + directoryName);
        }

    }
}
