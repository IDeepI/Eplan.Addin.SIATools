using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace WireMarking
{
    public static class ExportToExcel
    {
        // Excel variables
        private static Application xlApp;
        private static Workbook xlWorkBook;
        private static Worksheet xlWorkSheet1;
        private static Worksheet xlWorkSheet2;
        private static object misValue = System.Reflection.Missing.Value;

        // Regexp var
        private static string pattern = @"[^А-яЁё]+";
        private static string target = "";
        private static Regex regex = new Regex(pattern);

        // Var common for methods
        private static int rowNumberVO32 = 1;
        private static int rowNumberVO40 = 1;
        private static int rowNumberVO45 = 1;
        private static int rowNumberEmpty = 1;
        private static int rowNumber = 1;

        // Total Mark type count (30 mm each)
        private static int VO32 = 0;
        private static int VO40 = 0;
        private static int VO45 = 0;
        private static int noMark = 0;

        // Collumn count
        private static int columnNumber = 1;
        private static string tmpMarkType = "Not defined";
        // First section sheet
        private static string[,] sheetArray1 = null;
        // Second section sheet
        private static string[,] sheetArray2 = null;

        public static void Execute(List<EplanLabellingDocumentPageLine> listOfLines, string xlsFileName, Eplan.EplApi.Base.Progress progress)
        {
            Application xlApp = new Application();
            sheetArray1 = new string[listOfLines.Count * 2, 10];
            sheetArray2 = new string[listOfLines.Count * 2, 10];

            try
            {
                if (xlApp == null)
                {
                    DoWireMarking.DoWireMarking.MassageHandler("Excel is not properly installed!!");
                    return;
                }

                xlWorkBook = xlApp.Workbooks.Add(misValue);

                // Sheet count
                int sheetNumber = 1;
                xlWorkSheet1 = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetNumber);
                // Add as last
                xlWorkBook.Worksheets.Add(After: xlWorkSheet1);
                xlWorkSheet2 = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetNumber + 1);

                /// Name of RMU
                string boxName;

                for (int i = 0; i < listOfLines.Count; i++)
                {
                    boxName = listOfLines[i].Label?.Property[1]?.PropertyValue;

                    progress.BeginPart(40.0 / listOfLines.Count, "Write : " + boxName);
                    // Control new sheet creation
                    sheetNumber = ManageSheets(listOfLines, sheetNumber, boxName, i);
                    
                    // Total Mark type count 
                    VO32 += rowNumberVO32 - 1;
                    VO40 += rowNumberVO40 - 1;
                    VO45 += rowNumberVO45 - 1;
                    noMark += rowNumberEmpty - 1;

                    // Select column for each type of mark
                    SelectMarkType(listOfLines, ref columnNumber, ref tmpMarkType, ref rowNumber, i);
                    // Write marking name into arrays
                    WriteDataInCells(sheetArray1, listOfLines, columnNumber, rowNumber, i, "1");
                    WriteDataInCells(sheetArray2, listOfLines, columnNumber, rowNumber, i, "2");
                    rowNumber += 2;

                    progress.EndPart();
                }

                // Write array on sheet
                WriteArray<string>(xlWorkSheet1, 1, 1, sheetArray1);
                WriteArray<string>(xlWorkSheet2, 1, 1, sheetArray2);

                DoWireMarking.DoWireMarking.MassageHandler("VO32 - " + VO32 + "\nVO40 - " + VO40 + "\nVO45 - " + VO45 + "\nnoMark - " + noMark);

                xlWorkBook.SaveAs(xlsFileName, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, XlSaveConflictResolution.xlLocalSessionChanges, misValue, misValue, misValue, misValue);

                Debug.WriteLine($"Excel file created , you can find it in: \"{xlsFileName}\"");
            }
            catch (Exception ex)
            {
                DoWireMarking.DoWireMarking.ErrorHandler("ExportToExcel" + "\nlistOfLines.Count " + listOfLines.Count + "\nrowNumber " + rowNumber + "\ncolumnNumber " + columnNumber, ex);
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
        /// <summary>
        /// Write array on Excel sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"> Object </param>
        /// <param name="startRow"></param>
        /// <param name="startColumn"></param>
        /// <param name="array"></param>
        private static void WriteArray<T>(this _Worksheet sheet, int startRow, int startColumn, T[,] array)
        {
            var row = array.GetLength(0);
            var col = array.GetLength(1);
            Range c1 = (Range)sheet.Cells[startRow, startColumn];
            Range c2 = (Range)sheet.Cells[startRow + row - 1, startColumn + col - 1];
            Range range = sheet.Range[c1, c2];
            range.Value = array;
        }
        /// <summary>
        /// Write marking data into array
        /// </summary>
        /// <param name="sheetArray"> 2D Array </param>
        /// <param name="listOfLines"></param>
        /// <param name="columnNumber"></param>
        /// <param name="rowNumber"></param>
        /// <param name="i"></param>
        /// <param name="section"></param>
        private static void WriteDataInCells(string[,] sheetArray, List<EplanLabellingDocumentPageLine> listOfLines, int columnNumber, int rowNumber, int i, string section)
        {
            sheetArray[rowNumber - 1, columnNumber - 1] = tmpMarkType;
            sheetArray[rowNumber, columnNumber - 1] = tmpMarkType;

            string wireName = listOfLines[i].Label?.Property[9]?.PropertyValue.Replace("#", section).Replace("*", "");

            sheetArray[rowNumber - 1, columnNumber] = wireName;
            sheetArray[rowNumber, columnNumber] = wireName;
        }
        /// <summary>
        /// Control new sheet creation
        /// </summary>
        /// <param name="listOfLines"></param>
        /// <param name="sheetNumber"></param>
        /// <param name="boxName"></param> 
        /// <param name="i"> Count of data in object list </param>
        /// <returns></returns>
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
                rowNumberVO45 = 1;
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
        /// <summary>
        /// Saving row count for old mark type and selecting new type of mark
        /// </summary>
        /// <param name="listOfLines"></param>
        /// <param name="columnNumber"></param>
        /// <param name="tmpMarkType"></param>
        /// <param name="rowNumber"></param>
        /// <param name="i"></param>
        private static void SelectMarkType(List<EplanLabellingDocumentPageLine> listOfLines, ref int columnNumber, ref string tmpMarkType, ref int rowNumber, int i)
        {
            if (tmpMarkType != listOfLines[i].Label?.Property[3]?.PropertyValue)
            {
                // Save row count
                switch (tmpMarkType)
                {
                    case "VO-032BN4":
                        rowNumberVO32 = rowNumber;
                        break;

                    case "VO-040BN4":

                        rowNumberVO40 = rowNumber;
                        break;

                    case "VO-045BN4":

                        rowNumberVO45 = rowNumber;
                        break;
                    default:

                        rowNumberEmpty = rowNumber;
                        break;
                }


                tmpMarkType = listOfLines[i].Label?.Property[3]?.PropertyValue;
                // Change row count
                switch (tmpMarkType)
                {
                    case "VO-032BN4":
                        columnNumber = 1;
                        rowNumber = rowNumberVO32;
                        break;

                    case "VO-040BN4":
                        columnNumber = 3;
                        rowNumber = rowNumberVO40;
                        break;

                    case "VO-045BN4":
                        columnNumber = 5;
                        rowNumber = rowNumberVO45;
                        break;

                    default:
                        columnNumber = 7;
                        rowNumber = rowNumberEmpty;
                        break;
                }
            }
        }
        /// <summary>
        /// Creating new excel book sheet
        /// </summary>
        /// <param name="xlWorkSheet"></param>
        /// <param name="boxName"></param>
        /// <param name="curentSection"></param>
        private static void CreateBoxSheet(Worksheet xlWorkSheet, string boxName, int curentSection)
        {
            xlWorkSheet.Name = regex.Replace(boxName, target).Trim() + " секция " + curentSection;
        }

    }
}
