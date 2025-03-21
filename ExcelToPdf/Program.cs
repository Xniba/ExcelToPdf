﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
            string directoryPath = GetPathToDirectory("Files");
            string baseDirectoryPath = directoryPath + @"\BaseFiles";
            string newDirectoryPath = directoryPath + @"\NewFiles";

            CreateNewDirectory(newDirectoryPath);

            string[] excelList = ReturnPathForFilesWithExtensionFromDirectory(baseDirectoryPath, ".xlsx");
            string[] pdfList = PreparePathForFilesInNewDirectory(excelList, newDirectoryPath, ".pdf");

            ExcelToPdf(excelList, pdfList);

            Console.WriteLine("\n"+"All finished correctly, to close app press any key");
            Console.ReadKey();

            return;
        }
        static string[] ReturnPathForFilesWithExtensionFromDirectory(string directoryPath, string fileExtension)
        {
            Queue<string> baseFilesQueue = new Queue<string>();
            string[] baseFilesList = new string[0];

            if ('.' != fileExtension[0])
            {
                fileExtension = '.' + fileExtension;
            }

            try
            {
                baseFilesList = Directory.GetFiles(directoryPath, "*"+fileExtension);
                if (0 == baseFilesList.Length)
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
            for (int i = 0; i < baseFilesList.Length; i++)
            {
                if ('~' != baseFilesList[i].Substring(baseFilesList[i].LastIndexOf(@"\")+1 )[0])
                {
                    baseFilesQueue.Enqueue(baseFilesList[i]);
                }
                else
                {
                    temporaryFileExist = true;
                }
            }

            if (temporaryFileExist)
            {
                return baseFilesQueue.ToArray();
            }
            else
            {
                return baseFilesList;
            }
        }
        static string[] PreparePathForFilesInNewDirectory(string[] baseList, string newDirectoryPath)
        {
            string[] newList = new string[baseList.Length];
            try
            {
                string fileExtension = baseList[0].Substring(baseList[0].LastIndexOf(@"."));
                for (int i = 0; i < baseList.Length; i++)
                {
                    newList[i] = newDirectoryPath + baseList[i].Substring(baseList[i].LastIndexOf(@"\"));
                }
            }
            catch
            {
                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }

            return newList;
        }
        static string[] PreparePathForFilesInNewDirectory(string[] baseList, string newDirectoryPath, string newFileExtension)
        {
            if ('.' != newFileExtension[0])
            {
                newFileExtension = '.' + newFileExtension;
            }

            string[] newList = new string[baseList.Length];
            try
            {
                string fileExtension = baseList[0].Substring(baseList[0].LastIndexOf(@"."));
                for (int i = 0; i < baseList.Length; i++)
                {
                    newList[i] = newDirectoryPath +  
                    baseList[i]
                    .Substring(baseList[i].LastIndexOf(@"\"))
                    .Replace(fileExtension, newFileExtension);
                }
            }
            catch
            {
                Console.WriteLine("\nPlease contact the IT department");
                CloseApp();
            }

            return newList;
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

                Console.WriteLine("File conversion from .xlsx to .pdf has started\n");
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

                Console.WriteLine("\nAll files converted and saved in directory:");
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

            path = path.Substring(0, ile) + @"\" + directoryName;
            return path;
        }
    }
}
