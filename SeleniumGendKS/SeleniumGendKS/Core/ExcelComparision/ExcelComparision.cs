using Microsoft.Office.Interop.Excel;
using System.IO;
using static System.Windows.Forms.DataFormats;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace SeleniumGendKS.Core.ExcelComparision
{
    internal static class ExcelComparision
    {
        //private static readonly Application excelFile_1 = new Application();
        //private static readonly Application excelFile_2 = new Application();
        #region Compare Method
        public static bool IsFilesEqual(string path1, string path2)
        {
            bool isEqual = true;
            Application excelFile_1 = new Application();
            Application excelFile_2 = new Application();
            Workbook xbook_1 = excelFile_1.Workbooks.Open(path1);
            Workbook xbook_2 = excelFile_2.Workbooks.Open(path2);

            if (xbook_1.Worksheets.Count != xbook_2.Worksheets.Count)
            {
                isEqual = false;
            }
            else
            {
                for (int i = 0; i < xbook_1.Worksheets.Count; i++)
                {
                    Worksheet xsheet_1 = (Worksheet)xbook_1.Sheets[i + 1];
                    Range xrange_1 = xsheet_1.UsedRange;

                    Worksheet xsheet_2 = (Worksheet)xbook_2.Sheets[i + 1];
                    Range xrange_2 = xsheet_2.UsedRange;

                    for (int j = 0; j < xrange_1.Rows.Count; j++)
                    {
                        for (int k = 0; k < xrange_1.Columns.Count; k++)
                        {
                            Range cell_1 = (Range)xrange_1.Cells[j + 1, k + 1];
                            Range cell_2 = (Range)xrange_2.Cells[j + 1, k + 1];
                            if (cell_1.Value2 != cell_2.Value2)
                            {
                                isEqual = false;
                                Console.WriteLine("Compare between the generated file and baseline excel file: Fail\n");
                                Console.WriteLine("("+ (cell_1.Address) +")=" + cell_1.Value2 + " vs " + "(" + (cell_2.Address) + "-Baseline)=" + cell_2.Value2 + "\n");
                                break;
                            }
                        }
                        if (!isEqual) // isEqual == false
                        {
                            break;
                        }
                    }
                }
            }
            xbook_1.Close();
            xbook_2.Close();
            return isEqual;
        }

        public static bool IsSheetsEqual(string path1, string path2, string sheetName, string? fileName=null)
        {
            bool isEqual = true;
            Application excelFile_1 = new Application();
            Application excelFile_2 = new Application();
            Workbook xbook_1 = excelFile_1.Workbooks.Open(path1);
            Workbook xbook_2 = excelFile_2.Workbooks.Open(path2);
            var worksheet1 = xbook_1.Worksheets[sheetName];
            var worksheet2 = xbook_2.Worksheets[sheetName];
            Range xrange_1 = worksheet1.UsedRange;
            Range xrange_2 = worksheet2.UsedRange;

            for (int i= 0; i < xrange_1.Rows.Count; i++)
            {
                for (int j= 0; j < xrange_1.Columns.Count; j++)
                {
                    Range cell_1 = (Range)xrange_1.Cells[i + 1, j + 1];
                    Range cell_2 = (Range)xrange_2.Cells[i + 1, j + 1];
                    if (cell_1.Value2 != cell_2.Value2)
                    {
                        isEqual = false;

                        // Writing the difference into a txt file
                        string fileDiffTxt = "excelcomparediff.txt", summaryTCDes, summaryTCVal;
                        FileStream fs = new FileStream(Path.GetFullPath(@"../../../../TestResults/") + fileDiffTxt, FileMode.Append);
                        Console.WriteLine(summaryTCDes = "Compare between the generated file and baseline excel file: '" + fileName + "'-Sheet-'" + sheetName + "' Fail\n");
                        Console.WriteLine(summaryTCVal = "(" + (cell_1.Address) + ")=" + cell_1.Value2 + " vs " + "(" + (cell_2.Address) + "-Baseline)=" + cell_2.Value2 + "\n");
                        fs.Close();
                        File.AppendAllText(Path.GetFullPath(@"../../../../TestResults/") + fileDiffTxt, summaryTCDes + summaryTCVal + Environment.NewLine);

                        // Save as (Overwrite existing file) the downloaded filename and put it in the 'TestResult' folder
                        excelFile_1.DisplayAlerts = false;
                        xbook_1.SaveAs(Filename: Path.GetFullPath(@"../../../../TestResults/" + fileName.Replace("Book", "New")), AccessMode: XlSaveAsAccessMode.xlNoChange);
                        break;
                    }
                }
                if (!isEqual)
                {
                    break;
                }
            }
            xbook_1.Close();
            xbook_2.Close();
            return isEqual;
        }

        public static bool IsTheFormatSheetsEqual(string path1, string path2, string sheetName, string? fileName = null)
        {
            bool isEqual = true;
            Application excelFile_1 = new Application();
            Application excelFile_2 = new Application();
            Workbook xbook_1 = excelFile_1.Workbooks.Open(path1);
            Workbook xbook_2 = excelFile_2.Workbooks.Open(path2);
            var worksheet1 = xbook_1.Worksheets[sheetName];
            var worksheet2 = xbook_2.Worksheets[sheetName];
            Range xrange_1 = worksheet1.UsedRange;
            Range xrange_2 = worksheet2.UsedRange;

            for (int i = 0; i < xrange_1.Rows.Count; i++)
            {
                for (int j = 0; j < xrange_1.Columns.Count; j++)
                {
                    Range cell_1 = (Range)xrange_1.Cells[i + 1, j + 1];
                    Range cell_2 = (Range)xrange_2.Cells[i + 1, j + 1];
                    if (cell_1.Borders.LineStyle.ToString() != cell_2.Borders.LineStyle.ToString() ||
                        cell_1.Font.Name.ToString() != cell_2.Font.Name.ToString() ||
                        cell_1.Font.Color.ToString() != cell_2.Font.Color.ToString() ||
                        cell_1.NumberFormat.ToString() != cell_2.NumberFormat.ToString() ||
                        cell_1.Interior.Color.ToString() != cell_2.Interior.Color.ToString()) // --> Compare the fill color
                    {
                        isEqual = false;

                        // Writing the difference into a txt file
                        string fileDiffTxt = "excelcomparediff.txt", summaryTCDes, summaryTCVal;
                        FileStream fs = new FileStream(Path.GetFullPath(@"../../../../TestResults/") + fileDiffTxt, FileMode.Append);
                        Console.WriteLine(summaryTCDes = "Compare (Format) between the generated file and baseline excel file: '" + fileName + "'-Sheet-'" + sheetName + "' Fail\n");
                        Console.WriteLine(summaryTCVal = "(" + (cell_1.Address) + ")-Borders=" + cell_1.Borders.LineStyle + " vs " + "(" + (cell_2.Address) + "-Baseline)-Borders=" + cell_2.Borders.LineStyle + "\n" +
                                                         "(" + (cell_1.Address) + ")-FontName=" + cell_1.Font.Name + " vs " + "(" + (cell_2.Address) + "-Baseline)-FontName=" + cell_2.Font.Name + "\n" +
                                                         "(" + (cell_1.Address) + ")-FontColor=" + cell_1.Font.Color + " vs " + "(" + (cell_2.Address) + "-Baseline)-FontColor=" + cell_2.Font.Color + "\n" +
                                                         "(" + (cell_1.Address) + ")-NumberFormat=" + cell_1.NumberFormat + " vs " + "(" + (cell_2.Address) + "-Baseline)-NumberFormat=" + cell_2.NumberFormat + "\n" +
                                                         "(" + (cell_1.Address) + ")-FilledColor=" + cell_1.Interior.Color + " vs " + "(" + (cell_2.Address) + "-Baseline)-FilledColor=" + cell_2.Interior.Color + "\n");             
                        fs.Close();
                        File.AppendAllText(Path.GetFullPath(@"../../../../TestResults/") + fileDiffTxt, summaryTCDes + summaryTCVal + Environment.NewLine);

                        // Save as (Overwrite existing file) the downloaded filename and put it in the 'TestResult' folder
                        excelFile_1.DisplayAlerts = false;
                        xbook_1.SaveAs(Filename: Path.GetFullPath(@"../../../../TestResults/" + fileName.Replace("Book", "New")), AccessMode: XlSaveAsAccessMode.xlNoChange);
                        break;
                    }
                }
                if (!isEqual)
                {
                    break;
                }
            }
            xbook_1.Close();
            xbook_2.Close();
            return isEqual;
        }
        #endregion

        #region Common
        public static void MultiplyPercentAndGet2Digits(string path, string sheetName, string range)
        {
            Application excelFile = new Application();
            Workbook xbook = excelFile.Workbooks.Open(path);
            var worksheet = xbook.Worksheets[sheetName];
            Range xrange = worksheet.Range(range); // --> or Range[cell1, cell2];
            for (int j = 0; j < xrange.Rows.Count; j++)
            {
                for (int k = 0; k < xrange.Columns.Count; k++)
                {
                    Range cell = (Range)xrange.Cells[j + 1, k + 1];
                    if (cell.Value2 == null)
                    {
                        cell.Value = cell.Value;
                    }
                    else 
                    {
                        cell.Value = Math.Round(cell.Value * 100, 2, MidpointRounding.AwayFromZero);
                        cell.NumberFormat = "#,##0.00"; // --> org format = @"0.0%_);(0.0%)";
                    }
                }
            }

            // Save Workbook
            xbook.Save();

            // Close Excel files
            xbook.Close();
        }

        public static void MultiplyPercentAndGetXDigits(string path, string sheetName, string range, int roundXDigits, string? format=null)
        {
            Application excelFile = new Application();
            Workbook xbook = excelFile.Workbooks.Open(path);
            var worksheet = xbook.Worksheets[sheetName];
            Range xrange = worksheet.Range(range); // --> or Range[cell1, cell2];
            for (int j = 0; j < xrange.Rows.Count; j++)
            {
                for (int k = 0; k < xrange.Columns.Count; k++)
                {
                    Range cell = (Range)xrange.Cells[j + 1, k + 1];
                    if (cell.Value2 != null)
                    {
                        cell.Value = Math.Round(cell.Value * 100, roundXDigits, MidpointRounding.AwayFromZero);
                        cell.NumberFormat = format; // --> org format = @"0.0%_);(0.0%)"; // 2digits = "#,##0.00"
                    }
                    else
                    {
                        cell.Value = cell.Value;
                    }
                }
            }

            // Save Workbook
            xbook.Save();

            // Close Excel files
            xbook.Close();
        }

        public static void SortColumnRange(string path, string sheetName, string range)
        {
            Application excelFile = new Application();
            Workbook xbook = excelFile.Workbooks.Open(path);
            var worksheet = xbook.Worksheets[sheetName];
            Range xrange = worksheet.Range(range);

            // Sort
            xrange.Sort(xrange.Columns[1],XlSortOrder.xlAscending);

            // Save Workbook
            xbook.Save();

            // Close Excel files
            xbook.Close();
        }

        public static void ReOrderColumns(string path, string sheetName, string copyColumn, string insertColumn)
        {
            // Get path and Sheetname
            Application excelFile = new Application();
            Workbook xbook = excelFile.Workbooks.Open(path);
            var worksheet = xbook.Worksheets[sheetName];
            
            // Re order columns
            Range copyRange = worksheet.Range(copyColumn); // ex: copyColumn = "C:C";
            Range insertRange = worksheet.Range(insertColumn); // ex: insertColumn = "A:A";
            insertRange.Insert(XlInsertShiftDirection.xlShiftToRight, copyRange.Cut());

            // Save Workbook
            xbook.Save();

            // Close Excel files
            xbook.Close();
        }

        public static void ReOrderColumnsWithConditions(string path, string sheetName, string cellAddress, string columnName,string copyColumn, string insertColumn)
        {
            Application excelFile = new Application();
            Workbook xbook = excelFile.Workbooks.Open(path);
            var worksheet = xbook.Worksheets[sheetName];
            Range xrange = worksheet.Range(cellAddress);
            if (xrange.Text == columnName) // old: Value2
            {
                // Re order columns
                Range copyRange = worksheet.Range(copyColumn); // ex: copyColumn = "C:C";
                Range insertRange = worksheet.Range(insertColumn); // ex: insertColumn = "A:A";
                insertRange.Insert(XlInsertShiftDirection.xlShiftToRight, copyRange.Cut());
            }
            
            // Save Workbook
            xbook.Save();

            // Close Excel files
            xbook.Close();
        }
        #endregion
    }
}
