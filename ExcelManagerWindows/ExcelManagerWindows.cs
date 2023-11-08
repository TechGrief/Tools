using Microsoft.Office.Interop.Excel;
using System.Data.Common;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing.Printing;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.CodeDom;
using System.Diagnostics.SymbolStore;
using System.Xml.Linq;
using System.Xml;

namespace ExcelManagerWindows//com- > Microsoft Excel 16.0 Object library
{
    /// <summary>
    /// Create a new ExcelManagerWindows and do fun things with it!
    /// The PC is requiered to have Microsoft Excel installed!
    /// Printer configuration with Source Code...
    /// Time to make: ~ 1 day with classes - 8 nov 2023 - ERSTER TAG DER SEMESTEREXAMEN #3°B
    /// </summary>
    public class ExcelManagerWindows : IDisposable
    {
        private string PathToSource = "";
        private Excel.Application xlApp = null;
        private Excel.Workbook xlWorkBook = null;
        private List<Excel.Worksheet> xlWorkSheets = new List<Excel.Worksheet>();
        private object misValue = System.Reflection.Missing.Value;

        /// <summary>
        /// Create a new ExcelManagerWindows and do fun things with it!
        /// </summary>
        /// <param name="data">Pass the path or bytes (byte[]) of a xlsx Excel File</param>
        public ExcelManagerWindows(object data) 
        {
            if (data is byte[])
            {
                PathToSource = "ExcelManagerWindows_tmp_file_" + new Random().Next(1000) + "_" + new Random().Next(1000) + ".xlsx";
                File.WriteAllBytes(PathToSource, data as byte[]);
            }
            else PathToSource = data as string;

            if (!File.Exists(PathToSource)) return;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(PathToSource);
            foreach (Excel.Worksheet item in xlWorkBook.Worksheets)
                xlWorkSheets.Add(item);
        }


        /// <summary>
        /// Edit a specific Cell in a selected sheet.
        /// </summary>
        /// <param name="column">Set the Cell column. Example: D</param>
        /// <param name="row">Set the Cell row. Example: 13</param>
        /// <param name="newValue">Set the new Cell value. Example: Test</param>
        public bool UpdateCell(string column, int row, dynamic newValue, int sheet = 1)
        {
            if (sheet <= xlWorkSheets.Count)
            {
                //Attribute a value to a selected cell in a selected sheet
                var cell = (Excel.Range)xlWorkSheets[sheet-1].Cells[row, column];
                cell.Value2 = newValue;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Export or Save your Sheet/s as PDF or xlsx Excel File.
        /// Returns (bool) if path is set.
        /// Returns (string) with tmp-path to pdf if path is null
        /// </summary>
        /// <param name="path">Set Path to Target PDF.</param>
        /// <param name="sheet">Select sheet to export or 0 for All.</param>
        /// <param name="format">Select file type, can be pdf or xlsx.</param>
        public object ExportOrSave(string? path = null, int sheet = 0, string format = "pdf")
        {
            if (format != "pdf" && format != "xlsx") format = "pdf";

            bool returnPath = false;
            try
            {
                if (path == null)
                {
                    path = Path.Combine(System.IO.Path.GetTempPath(), "ExcelManagerWindows_tmp_file_" + new Random().Next(1000) + "_" + new Random().Next(1000) + "." + format);
                    returnPath = true;
                }

                if (sheet == 0)
                {
                    if(format == "pdf")
                        xlWorkBook.ExportAsFixedFormat2(XlFixedFormatType.xlTypePDF, path);
                    if (format == "xlsx")
                        xlWorkBook.SaveAs2(path);
                }
                else if (sheet > 0 && sheet <= xlWorkSheets.Count)
                {
                    if (format == "pdf")
                        xlWorkSheets[sheet - 1].ExportAsFixedFormat2(XlFixedFormatType.xlTypePDF, path);
                    if (format == "xlsx")
                        xlWorkSheets[sheet - 1].SaveAs2(path);
                }
                else return (returnPath == false ? false : "");

                return (returnPath == false ? true : path);
            }
            catch (Exception) { }

            return (returnPath == false ? false : "");
        }

        public void Dispose()
        {
            try { xlWorkSheets.Clear(); } catch (Exception) { }
            try { GC.Collect(); } catch (Exception) { }
            try { GC.WaitForPendingFinalizers(); } catch (Exception) { }
            try { if (xlWorkBook != null) xlWorkBook.Close(false, Type.Missing, Type.Missing); } catch (Exception) { }
            try { Marshal.FinalReleaseComObject(xlWorkBook); } catch (Exception) { }
            try { if(xlApp != null) xlApp.Quit(); } catch (Exception) { }
            try { Marshal.FinalReleaseComObject(xlApp); } catch (Exception) { }
        }
    }
}
