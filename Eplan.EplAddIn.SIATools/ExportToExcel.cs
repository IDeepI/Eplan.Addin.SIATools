using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace SIATools
{
    public static class ExportToExcel
    {
        // Excel variables
        private static Application xlApp;
        private static Workbook xlWorkBook;
        private static Worksheet xlWorkSheet;

        private static object misValue = System.Reflection.Missing.Value;

        // Regexp var
        private static string pattern = @"[^А-яЁё]+";
        private static string target = "";
        private static Regex regex = new Regex(pattern);

        // Var common for methods
        private static List<string> markType = new List<string>() { };
        private static Dictionary<string, int> markTypeRow = new Dictionary<string, int>();

        private static int rowNumber = 1;

        // Sheet count
        private static int xlsMainSheetCounter = 1;

        // Collumn count
        private static int columnNumber = 1;
        private static string tmpMarkType = "Not defined";
        // First section sheet
        private static string[,] sheetArray = null;

        /// Name of RMU
        private static string boxName;

        public static void Execute(List<EplanLabellingDocumentPageLine> listOfLines, string xlsFileName, Eplan.EplApi.Base.Progress progress)
        {
            markType = new List<string>() {
                "ПуГВнг(А)-LS 1х1",
                "ПуГВнг(А)-LS 1х1,5",
                "ПуГВнг(А)-LS 1х2,5",
                "ПуГВнг(А)-LS 1х2,5 Ж-З",
                "ПуГВнг(А)-LS 1х4",
                "ПуГВнг(А)-LS 1х4 Ж-З",
                ""
            };

            int sectionCount = 2;

            foreach (var mark in markType)
            {
                markTypeRow[mark] = 1;
            }  

            Application xlApp = new Application();
            sheetArray = new string[listOfLines.Count * 2, markType.Count * 2];
            for (int i = 0; i < markType.Count; i++)
            {
                sheetArray[0, i * 2] = markType[i];
            }

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
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetNumber);

                for (int j = 1; j <= sectionCount; j++)
                {
                    rowNumber = 1;
                    for (int i = 0; i < listOfLines.Count; i++)
                    {
                        boxName = listOfLines[i].Label?.Property[1]?.PropertyValue;

                        if (boxName.Substring(0,4) == "ИВЕЖ")
                        {                        
                            progress.BeginPart(20.0 / listOfLines.Count, "Writing : " + boxName + " секция "+ j);
                            // Control new sheet creation
                            sheetNumber = ManageSheets(listOfLines, sheetNumber, boxName, i, j);
                            // Select column for each type of mark
                            SelectMarkType(listOfLines, ref columnNumber, ref tmpMarkType, ref rowNumber, i);
                            // Write marking name into arrays
                            WriteDataInCells(sheetArray, listOfLines, columnNumber, rowNumber, i, j, j == 1 ? 2 : 1);

                            rowNumber += 2;

                            progress.EndPart();

                            if (progress.Canceled())
                            {
                                progress.EndPart(true);
                                i = listOfLines.Count;
                            }
                        }
                        else if (j == 1)
                        {
                            progress.BeginPart(20.0 / listOfLines.Count, "Writing : " + boxName + " секция " + j);
                            // Control new sheet creation
                            sheetNumber = ManageSheets(listOfLines, sheetNumber, boxName, i, j);
                            // Select column for each type of mark
                            SelectMarkType(listOfLines, ref columnNumber, ref tmpMarkType, ref rowNumber, i);
                            // Write marking name into arrays
                            CableDataInCells(sheetArray, listOfLines, columnNumber, rowNumber, i, j, j == 1 ? 2 : 1);

                            rowNumber += 2;

                            progress.EndPart();

                            if (progress.Canceled())
                            {
                                progress.EndPart(true);
                                i = listOfLines.Count;
                            }
                        }
                    }
                }
                // Write array on sheet
                WriteArray<string>(xlWorkSheet, 1, 1, sheetArray);               

                xlWorkBook.SaveAs(xlsFileName, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, XlSaveConflictResolution.xlLocalSessionChanges, misValue, misValue, misValue, misValue);

                Debug.WriteLine($"Excel file created , you can find it in: \"{xlsFileName}\"");
            }
            catch (Exception ex)
            {
                DoWireMarking.DoWireMarking.ErrorHandler("ExportToExcel" + "\nlistOfLines.Count " + listOfLines.Count + "\nrowNumber " + rowNumber + "\ncolumnNumber " + columnNumber + "\nboxName " + boxName, ex);
                return;
            }
            finally
            {
                xlWorkBook?.Close(true, misValue, misValue);
                xlApp?.Quit();

                xlsMainSheetCounter = 1;
                Marshal.ReleaseComObject(xlWorkSheet);                
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
            sheet.Columns.AutoFit();
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
        private static void WriteDataInCells(string[,] sheetArray, List<EplanLabellingDocumentPageLine> listOfLines, int columnNumber, int rowNumber, int i, int section, int oppositeSection)
        {
            sheetArray[rowNumber - 1, columnNumber - 1] = tmpMarkType;
            sheetArray[rowNumber, columnNumber - 1] = tmpMarkType;

            string wireName = listOfLines[i].Label?.Property[9]?.PropertyValue.Replace("#", section.ToString()).Replace("^", oppositeSection.ToString()).Replace("*", "");            

            sheetArray[rowNumber - 1, columnNumber] =  wireName;
            sheetArray[rowNumber, columnNumber] =  wireName;
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
        private static void CableDataInCells(string[,] sheetArray, List<EplanLabellingDocumentPageLine> listOfLines, int columnNumber, int rowNumber, int i, int section, int oppositeSection)
        {
            sheetArray[rowNumber - 1, columnNumber - 1] = tmpMarkType;
            sheetArray[rowNumber, columnNumber - 1] = tmpMarkType;

            string cableName = listOfLines[i].Label?.Property[12]?.PropertyValue;
            string wireName = listOfLines[i].Label?.Property[9]?.PropertyValue.Replace("#", section.ToString()).Replace("^", oppositeSection.ToString()).Replace("*", "");

            sheetArray[rowNumber - 1, columnNumber] = cableName + "/" + wireName;
            sheetArray[rowNumber, columnNumber] = cableName + "/" + wireName;
        }
        /// <summary>
        /// Control new sheet creation
        /// </summary>
        /// <param name="listOfLines"></param>
        /// <param name="sheetNumber"></param>
        /// <param name="boxName"></param> 
        /// <param name="i"> Count of data in object list </param>
        /// <returns></returns>
        private static int ManageSheets(List<EplanLabellingDocumentPageLine> listOfLines, int sheetNumber, string boxName, int i, int section)
        {
            if (i == 0)
            {
                CreateBoxSheet(xlWorkSheet, boxName, section);                
            }
            else if (boxName == listOfLines[i - 1].Label?.Property[1]?.PropertyValue)
            {

            }
            else
            {
                // Write array on sheet
                WriteArray<string>(xlWorkSheet, 1, 1, sheetArray);
              
                // Clear Array
                sheetArray = new string[listOfLines.Count * 2, markType.Count * 2];
                for (int j = 0; j < markType.Count; j++)
                {
                    sheetArray[0, j * 2] = markType[j];
                }

                // Start row count from the begining
                foreach (var mark in markType)
                {
                    markTypeRow[mark] = 1;
                }
                rowNumber = 1;

                sheetNumber += 1;
                xlWorkBook.Worksheets.Add(After: xlWorkSheet);
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(sheetNumber); 

                CreateBoxSheet(xlWorkSheet, boxName, section);                
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
            if (tmpMarkType != listOfLines[i].Label?.Property[6]?.PropertyValue)
            {
                // Save row count
                markTypeRow[tmpMarkType] = rowNumber;

                tmpMarkType = listOfLines[i].Label?.Property[6]?.PropertyValue;                

                // Use last column if couldn't find markType
                if (!markType.Contains(tmpMarkType))
                {
                    tmpMarkType = "";
                }               

                // Change row count
                rowNumber = markTypeRow[tmpMarkType];

                columnNumber = markType.IndexOf(tmpMarkType) * 2 + 1;

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
            // string xlWorkSheetName = xlsSheetCounter + "." + regex.Replace(boxName, target).Trim() + " с." + curentSection;
            string xlWorkSheetName = xlsMainSheetCounter + "." + boxName.Trim() + "с" + curentSection;
            if (xlWorkSheetName.Length > 31)
            {
                xlWorkSheetName = xlWorkSheetName.Substring(0, 31);                
            }
            xlWorkSheet.Name = xlWorkSheetName;
            xlsMainSheetCounter++;
        }

    }
}
