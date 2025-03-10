using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

//NuGet
using Microsoft.Office.Interop.Excel;

namespace ExcelToPdf
{
    internal class Program
    {

        static void Test_ExcelToPdf()
        {
            string excelFile = "C:\\Users\\plradbad\\source\\repos\\Xniba\\ExcelToPdf\\ExcelToPdf\\Files\\ExcelFiles\\CA001 - ReplacementFile.xlsx";
            string pfdFile = "C:\\Users\\plradbad\\source\\repos\\Xniba\\ExcelToPdf\\ExcelToPdf\\Files\\PdfFiles\\CA001 - ReplacementFile.pdf";


            Application excelApp = new Application();
            Workbook workbook_1;
            Workbook workbook_2;
            Worksheet worksheet;

            // 1.1 Open workbook from file path
            workbook_1 = excelApp.Workbooks.Open(excelFile);

            // 1.2. Open from workbook one worksheet, thats one with we want to save 
            worksheet = (Worksheet)workbook_1.Sheets[2];

            // 1.3. Create second workboook from instance of excel
            workbook_2 = excelApp.Workbooks.Add();

            // 1.4. Copy worksheets from first workbook to secound wborkbook in first worksheets (wb2.Sheets[1])
            worksheet.Copy(workbook_2.Sheets[1]);

            // 1.5. Save new workbook as a PDF file
            workbook_2.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pfdFile);

            // 1.6. Close instance
            workbook_1.Close(0);
            workbook_2.Close(0);

            // 1.7. Info:

            Console.WriteLine("\n" + "All files converted, saved in directory:");
            Console.WriteLine("lokalizacja pliku");

            // 1.8. Close 
            excelApp.Quit();

            return;
        }
        static void Main(string[] args)
        {

            Test_ExcelToPdf();
            return;
            ////Parameters////
            string directoryPathFiles = GetPathToDirectory("Files");
            string directoryPathExcel = directoryPathFiles + @"\ExcelFiles";
            string newDirectoryPathPdf = directoryPathFiles + @"\PdfFiles";
 
            CreateNewDirectory(newDirectoryPathPdf);

            string[] excelList = ReadNameOfExcelFiles(directoryPathExcel);
            string[] pdfList = PreparePathForFilesToNewDirectoryWithChangedExtension(excelList, newDirectoryPathPdf, ".xlsx", ".pdf");

            ExcelToPdf(excelList, pdfList);

            Console.WriteLine("\n"+"All finished, to close app press any key");
            Console.ReadKey();

            return;
        }
        static string[] ReadNameOfExcelFiles(string directoryPath)
        {
            string[] excelNamesList = new string[0];
            try
            {
                excelNamesList = Directory.GetFiles(directoryPath, "*.xlsx");
            }
            catch 
            { 
                CloseApp();
            }

            Queue<string> excelNamesQueue = new Queue<string>();
            for (int i = 0; i < excelNamesList.Length; i++)
            { 
                if (excelNamesList[i][0] != '~')
                {
                    excelNamesQueue.Enqueue(excelNamesList[i]); 
                }
            }

            return excelNamesQueue.ToArray();
        }
        static string[] PreparePathForFilesToNewDirecotryFromExcelFilesToPdfFiles(string[] excelList, string newDirectoryPath)
        {
            string[] pdfList = new string[excelList.Length];

            try
            {
                for (int i = 0; i < excelList.Length; i++)
                {
                    pdfList[i] = newDirectoryPath +  
                    excelList[i]
                    .Substring(excelList[i].LastIndexOf(@"\"))
                    .Replace(".xlsx", ".pdf");
                    
                }
            }
            catch
            {
                CloseApp();
            }

            return pdfList;
        }
        static string[] PreparePathForFilesToNewDirectoryWithChangedExtension(string[] excelList, string newDirectoryPath, string orginalExtension, string newExtension)
        {
            string[] pdfList = new string[excelList.Length];

            try
            {
                for (int i = 0; i < excelList.Length; i++)
                {
                    pdfList[i] = newDirectoryPath +
                    excelList[i]
                    .Substring(excelList[i].LastIndexOf(@"\"))
                    .Replace(orginalExtension, newExtension);

                }
            }
            catch
            {
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

                Console.WriteLine("No permition to delete directory:\n" + newDirectoryPath);
                CloseApp();
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Can't delete or create new directory");
                Console.WriteLine($"Close all files from directory:\n" + newDirectoryPath);
                Console.ResetColor();

                CloseApp();
            }
        }
        static void ExcelToPdf(string[] excelList, string[] pdfList)
        {
            int amountOfFiles = excelList.Length;
            Console.WriteLine($"Amount found files: {amountOfFiles}");

            Application excelApp = new Application();
            Workbook workbook_1;
            Workbook workbook_2;
            Worksheet worksheet;

            Console.WriteLine("\nConverting files in progres"+ "\n");
            // 1. Create new pdf from excel
            for (int i = 0; i < amountOfFiles; i++)
            {
                // 1.1 Open workbook from file path
                workbook_1 = excelApp.Workbooks.Open(excelList[i]);

                // 1.2. Open from workbook one worksheet, thats one with we want to save 
                worksheet = (Worksheet)workbook_1.Sheets[2];

                // 1.3. Create second workboook from instance of excel
                workbook_2 = excelApp.Workbooks.Add();

                // 1.4. Copy worksheets from first workbook to secound wborkbook in first worksheets (wb2.Sheets[1])
                worksheet.Copy(workbook_2.Sheets[1]);

                // 1.5. Save new workbook as a PDF file
                workbook_2.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfList[i]);

                // 1.6. Close instance
                workbook_1.Close(0);
                workbook_2.Close(0);

                // 1.7. Info:
                Console.WriteLine("File {0} saved: {1}", i + 1, pdfList[i].Substring(pdfList[i].LastIndexOf(@"\") + 1, 6));
            }

            // 1.7. Info:

            Console.WriteLine("\n" + "All files converted, saved in directory:");
            Console.WriteLine(pdfList[0].Substring(0, pdfList[0].LastIndexOf(@"\")));

            // 1.8. Close 
            excelApp.Quit();

            return;
        }
        static void CloseApp([CallerMemberName] string methodName = "")
        {
            Debug.WriteLine($"!!!There was a problem in method: {methodName}");

            Console.WriteLine("\nPress Enter to close app");
            Console.ReadKey();
            Environment.Exit(0);
        }
        static string GetPathToDirectory()
        {
            string path = new DirectoryInfo(".").FullName;
            int ile = path.IndexOf("bin") - 1;
            if (ile < 0)
            {
                Console.WriteLine("Director not fount, please contact with your IT department");
                CloseApp();
            }
            else
            {
                path = path.Substring(0, ile);
            }
            return path;
        }
        static string GetPathToDirectory(string directoryName)
        {
            string path = new DirectoryInfo(".").FullName;
            int ile = path.IndexOf("bin") - 1;
            if (ile < 0)
            {
                Console.WriteLine("Director not fount, please contact with your IT department");
                CloseApp();
            }
            else
            {
                path = path.Substring(0, ile);
                path = path + @"\" + directoryName;
            }
            return path;
        }

    }
}
