using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using WireMarking;

namespace Eplan.Addin.WireMarking
{
    public static class ExportToExcel
    {
        private static Application xlApp;
        private static Workbook xlWorkBook;
        private static Worksheet xlWorkSheet1;
        private static Worksheet xlWorkSheet2;
        private static object misValue = System.Reflection.Missing.Value;

        private static string pattern = @"[^А-яЁё]+";
        private static string target = "";
        private static Regex regex = new Regex(pattern);

        private static int rowNumberVO32 = 1;
        private static int rowNumberVO40 = 1;
        private static int rowNumberVO48 = 1;
        private static int rowNumberEmpty = 1;
        private static int rowNumber = 1;
        private static string tmpMarkType = "Not defined";

        private static string[,] sheetArray1 = null;        
        private static string[,] sheetArray2 = null;        

        public static void Execute(List<EplanLabellingDocumentPageLine> listOfLines, string xlsFileName)
        {
            Application xlApp = new Application();
            sheetArray1 = new string[listOfLines.Count, 10];
            sheetArray2 = new string[listOfLines.Count, 10];

            try
            {
                if (xlApp == null)
                {
                    DoWireMarking.MassageHandler("Excel is not properly installed!!");
                    return;
                }

                xlWorkBook = xlApp.Workbooks.Add(misValue);

                // Sheet count
                int sheetNumber = 1;
                xlWorkSheet1 = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetNumber);
                // Add as last
                xlWorkBook.Worksheets.Add(After: xlWorkSheet1);
                xlWorkSheet2 = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetNumber + 1);
                // Collumn count
                int columnNumber = 1;

                string boxName;
                

                for (int i = 0; i < listOfLines.Count; i++)
                {
                    boxName = listOfLines[i].Label?.Property[1]?.PropertyValue;

                    sheetNumber = ManageSheets(listOfLines, sheetNumber, boxName, i);

                    SelectMarkType(listOfLines, ref columnNumber, ref tmpMarkType, ref rowNumber, i);

                    
                    WriteDataInCells(sheetArray1, listOfLines, columnNumber, rowNumber, i, "1");
                    WriteDataInCells(sheetArray2, listOfLines, columnNumber, rowNumber, i, "2");
                    rowNumber += 2;
                }

                // Write array on sheet
                WriteArray<string>(xlWorkSheet1, 1, 1, sheetArray1);
                WriteArray<string>(xlWorkSheet2, 1, 1, sheetArray2);

                xlWorkBook.SaveAs(xlsFileName, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, XlSaveConflictResolution.xlLocalSessionChanges, misValue, misValue, misValue, misValue);

                DoWireMarking.MassageHandler($"Excel file created , you can find it in: \"{xlsFileName}\"");
            }
            catch (Exception ex)
            {
                DoWireMarking.ErrorHandler("ExportToExcel", ex);
                return;
            }
            finally
            {
                xlWorkBook?.Close(true, misValue, misValue);
                xlApp?.Quit();

                Marshal.ReleaseComObject(xlWorkSheet1);
                Marshal.ReleaseComObject(xlWorkSheet2);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }
        private static void WriteArray<T>(this _Worksheet sheet, int startRow, int startColumn, T[,] array)
        {
            var row = array.GetLength(0);
            var col = array.GetLength(1);
            Range c1 = (Range)sheet.Cells[startRow, startColumn];
            Range c2 = (Range)sheet.Cells[startRow + row - 1, startColumn + col - 1];
            Range range = sheet.Range[c1, c2];
            range.Value = array;
        }

        private static void WriteDataInCells(string[,] sheetArray, List<EplanLabellingDocumentPageLine> listOfLines, int columnNumber, int rowNumber, int i, string section)
        {
            sheetArray[rowNumber - 1, columnNumber - 1] = tmpMarkType;
            sheetArray[rowNumber, columnNumber - 1] = tmpMarkType;

            string wireName = listOfLines[i].Label?.Property[9]?.PropertyValue.Replace("#", section).Replace("*", "");

            sheetArray[rowNumber - 1, columnNumber] = wireName;
            sheetArray[rowNumber, columnNumber] = wireName;

           /* sheetArray[columnNumber - 1, rowNumber - 1] = tmpMarkType;
            sheetArray[columnNumber - 1, rowNumber] = tmpMarkType;

            string wireName = listOfLines[i].Label?.Property[9]?.PropertyValue.Replace("#", section).Replace("*", "");

            sheetArray[columnNumber, rowNumber - 1] = wireName;
            sheetArray[columnNumber, rowNumber] = wireName;*/
         
           // xlWorkSheet.Cells[rowNumber, columnNumber] = listOfLines[i].Label?.Property[3]?.PropertyValue;
           // xlWorkSheet.Cells[rowNumber, columnNumber + 1] = listOfLines[i].Label?.Property[9]?.PropertyValue.Replace("#", section).Replace("*", "");
           // xlWorkSheet.Cells[rowNumber, columnNumber + 10] = listOfLines[i].Label?.Property[12]?.PropertyValue;

            /*Debug.Write($"{listOfLines[i].Label?.Property[1]?.PropertyValue}\t : \t");
            Debug.Write($"{section}\t : \t");
            Debug.Write($"{columnNumber}\t : \t");
            Debug.Write($"{rowNumber}\t : \t");
            Debug.Write($"{listOfLines[i].Label?.Property[3]?.PropertyValue}\t : \t");
            Debug.Write($"{listOfLines[i].Label?.Property[9]?.PropertyValue.Replace("#", section).Replace("*", "")}\t : \t");
            Debug.Write($"{listOfLines[i].Label?.Property[12]?.PropertyValue}\t : \t");

            Debug.WriteLine("");*/
        }

        private static int ManageSheets(List<EplanLabellingDocumentPageLine> listOfLines, int sheetNumber, string boxName, int i)
        {
            if (i == 0)
            {
                CreateBoxSheet(xlWorkSheet1, boxName, 1);
                CreateBoxSheet(xlWorkSheet2, boxName, 2);
            }
            else if (boxName == listOfLines[i - 1].Label?.Property[1]?.PropertyValue)
            {

            }
            else
            {
                // Write array on sheet
                WriteArray<string>(xlWorkSheet1, 1, 1, sheetArray1);
                WriteArray<string>(xlWorkSheet2, 1, 1, sheetArray2);
               
                // Clear Array
                sheetArray1 = new string[listOfLines.Count, 10];
                sheetArray2 = new string[listOfLines.Count, 10];

                // Start row count from the begining
                rowNumberVO32 = 1;
                rowNumberVO40 = 1;
                rowNumberVO48 = 1;
                rowNumberEmpty = 1;
                rowNumber = 1;

                sheetNumber += 2;
                xlWorkBook.Worksheets.Add(After: xlWorkSheet2);
                xlWorkSheet1 = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetNumber);
                xlWorkBook.Worksheets.Add(After: xlWorkSheet1);
                xlWorkSheet2 = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetNumber + 1);

                CreateBoxSheet(xlWorkSheet1, boxName, 1);
                CreateBoxSheet(xlWorkSheet2, boxName, 2);
            }

            return sheetNumber;
        }

        private static void SelectMarkType(List<EplanLabellingDocumentPageLine> listOfLines, ref int columnNumber, ref string tmpMarkType, ref int rowNumber, int i)
        {
            if (tmpMarkType != listOfLines[i].Label?.Property[3]?.PropertyValue)
            {
                // Save row count
                switch (tmpMarkType)
                {
                    case "VO-32":
                        rowNumberVO32 = rowNumber;
                        break;

                    case "VO-40":

                        rowNumberVO40 = rowNumber;
                        break;

                    case "VO-48":

                        rowNumberVO48 = rowNumber;
                        break;
                    default:

                        rowNumberEmpty = rowNumber;
                        break;
                }


                tmpMarkType = listOfLines[i].Label?.Property[3]?.PropertyValue;
                // Change row count
                switch (tmpMarkType)
                {
                    case "VO-32":
                        columnNumber = 1;
                        rowNumber = rowNumberVO32;
                        break;

                    case "VO-40":
                        columnNumber = 3;
                        rowNumber = rowNumberVO40;
                        break;

                    case "VO-48":
                        columnNumber = 5;
                        rowNumber = rowNumberVO48;
                        break;

                    default:
                        columnNumber = 7;
                        rowNumber = rowNumberEmpty;
                        break;
                }
            }
        }

        private static void CreateBoxSheet(Worksheet xlWorkSheet, string boxName, int curentSection)
        {

            xlWorkSheet.Name = regex.Replace(boxName, target).Trim() + " секция " + curentSection;
        }

    }
}
