using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

//dodać nugetsa
using Microsoft.Office.Interop.Excel;

namespace ExcelToPdf
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ////Parameters////
            string directoryPath = GetPathToDirectory() + @"\Files";              //existing direcroty
            string directoryPathExcel = directoryPath + @"\Excel";              //existing directory witch excel, to convert
            string newDirectoryPathPdf = directoryPath + @"\Pdf";                //new directory for PDFs
            string newDirectoryPathPdfSigned = directoryPath + @"\PdfSigned";    //new directory for signed PDFs 

            //List of names, excel and future pdf files
            string[] excelList;     //array of excel name and path
            string[] pdfList;       //array of pdf name and paths
            string[] pdfSignedList; //array of pdf name and paths

            // number of files to convert
            int amountOfFiles = 0;

            excelList = new string[] { @"C:\Users\Adrian\Desktop\Testy\Files\BaseFiles\CA001 - ReplacementFile.xlsx" };
            pdfList = new string[] { @"C:\Users\Adrian\Desktop\Testy\Files\BaseFiles\CA001 - ReplacementFile.pdf" };
            amountOfFiles = 1;
            ExcelToPdf(excelList, pdfList, amountOfFiles);
            return;

            ////Program////
            // 1. Create new directory for PDF
            CreateNewDirectory(newDirectoryPathPdf);

            // 2.1. Read excel files name from directory
            excelList = ReadNamesOfExcels(directoryPathExcel);

            // 2.1.1.
            amountOfFiles = excelList.Count();
            Console.WriteLine($"Amount found files: {amountOfFiles}");

            // 2.2. Prepare pdf file name (path)
            pdfList = NamesExcelsToPdf(excelList, newDirectoryPathPdf);

            // 4. Creating from excel new PDFs, in new directory 
            ExcelToPdf(excelList, pdfList, amountOfFiles);

            return;
        }
        static string[] ReadNamesOfExcels(string directoryPath)
        {
            // 0. Variable
            Queue<string> callerIds = new Queue<string>();  //Queue for later sort
            int i;                                          //variable for loop
            bool change = false;

            // 1. Read all files with ".xlsx" extension
            string[] list = Directory.GetFiles(directoryPath, "*.xlsx");

            // 2. Check all readed files, if ther is file started on "~" (thats mean instance of open file) don't save it in queue
            for (i = 0; i < list.Count(); i++)
            {
                //list[i][0]; - its mean, first (index 0) chart in string list[i]
                if (list[i][0] != '~')
                { callerIds.Enqueue(list[i]); } //Add to queue variable form array 
                else
                { change = true; }
            }

            // 3. If in step-2 find instance file, make below
            if (change)
            {
                list = new string[callerIds.Count()];   //resize array
                i = 0;                                  //reset loop counter
                foreach (var id in callerIds)           //Rewrite array
                {
                    list[i] = id;
                    i++;
                }
            }

            return list;
        }
        static string[] NamesExcelsToPdf(string[] excelList, string newDirectoryPath)
        {
            string[] pdfList = new string[excelList.Count()];

            // 1. Read names of Excel files and prepare names for PDF 
            int i = 0;
            foreach (string excel in excelList)
            {
                pdfList[i] = excelList[i].Substring(excelList[i].LastIndexOf(@"\"));
                pdfList[i] = newDirectoryPath + pdfList[i].Replace(".xlsx", ".pdf");
                i++;
            }
            return pdfList;
        }
        static void CreateNewDirectory(string newDirectoryPath)
        {
            //New directory for files
            try
            {
                if (Directory.Exists(newDirectoryPath))
                {
                    Directory.Delete(newDirectoryPath, true); //true, give permision to delete directory and all content
                    Directory.CreateDirectory(newDirectoryPath);
                }
                else
                {
                    Directory.CreateDirectory(newDirectoryPath);
                }
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Can't delete or create new directory");
                Console.WriteLine($"Close all files from directory: {newDirectoryPath}");
                Console.WriteLine($"Or give permission to write in: {newDirectoryPath.Substring(0, 35 - 10)}");
                Console.WriteLine("\nPress any key to close window");
                Console.ResetColor();
                Console.ReadKey();
                System.Environment.Exit(0);
            }
        }
        static void ExcelToPdf(string[] excelList, string[] pdfList, int amountOfFiles)
        {
            //Parameters
            //Create Excel App instance
            Application excelApp = new Application();
            Workbook wb;
            Workbook wb2;
            Worksheet ws;

            Console.WriteLine("\nConversion of files from .xlsx to .pdf has started");
            // 1. Create new pdf from excel
            for (int i = 0; i < amountOfFiles; i++)
            {
                // 1.1 Open workbook from file path
                wb = excelApp.Workbooks.Open(excelList[i]);

                // 1.2. Open from workbook one worksheet, thats one with we want to save 
                ws = (Worksheet)wb.Sheets[2]; //wb.Sheets[2].Name, return name of 2nd worksheet. Is counted from one(1)
                                              //ws = (Worksheet)wb.Sheets[2]; //////////////////////////try this

                // 1.3. Create second workboook from instance of excel
                wb2 = excelApp.Workbooks.Add();

                // 1.4. Copy worksheets from first workbook to secound wborkbook in first worksheets (wb2.Sheets[1])
                ws.Copy(wb2.Sheets[1]);

                // 1.5. Save new workbook as a PDF file
                wb2.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfList[i]);

                // 1.6. Close instance
                wb.Close(0);
                wb2.Close(0);

                // 1.7. Info:
                Console.WriteLine("File {0} saved: {1}", i + 1, pdfList[i].Substring(pdfList[i].LastIndexOf(@"\") + 1, 6));
            }

            // 1.7. Info:

            Console.WriteLine("All files converted, saved in directory:");
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
