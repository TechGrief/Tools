using OfficeOpenXml;

namespace ExcelManagerEPPlus
{
    /// <summary>
    /// .dlls requiered (min.):
    /// EPPlus.dll
    /// Microsoft.IO.RecyclableMemoryStream.dll
    /// Microsoft.Extensions.Configuration.Json.dll
    /// Microsoft.Extensions.Configuration.FileExtensions.dll
    /// Microsoft.Extensions.Configuration.dll
    /// Microsoft.Extensions.FileProviders.Physical.dll
    /// Microsoft.Extensions.FileProviders.Abstractions.dll
    /// Microsoft.Extensions.Configuration.Abstractions.dll
    /// Microsoft.Extensions.Primitives.dll
    /// </summary>
    public class ExcelManagerEPPlus : IDisposable
    {
        private ExcelPackage excelPackage = null;

        private string? PathToSource = "";

        /// <summary>
        /// Create a new ExcelManagerEPPlus and do fun things with it!
        /// </summary>
        /// <param name="data">Pass the path or bytes (byte[]) of a xlsx Excel File</param>
        public ExcelManagerEPPlus(object data)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            if (data is byte[])
            {
                PathToSource = "ExcelManagerWindows_tmp_file_" + new Random().Next(1000) + "_" + new Random().Next(1000) + ".xlsx";
                File.WriteAllBytes(PathToSource, data as byte[]);
            }
            else PathToSource = data as string;
            if (!File.Exists(PathToSource)) return;

            excelPackage = new ExcelPackage(PathToSource);
        }


        /// <summary>
        /// Edit a specific Cell in a selected sheet.
        /// </summary>
        /// <param name="column">Set the Cell column. Example: D</param>
        /// <param name="row">Set the Cell row. Example: 13</param>
        /// <param name="newValue">Set the new Cell value. Example: Test</param>
        public bool UpdateCell(string column, int row, object newValue, int sheet = 1)
        {
            if (sheet <= excelPackage.Workbook.Worksheets.Count)
            {
                //Attribute a value to a selected cell in a selected sheet
                excelPackage.Workbook.Worksheets[sheet - 1].Cells[column + row].Value = newValue;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Export or Save your Sheet/s as xlsx Excel File.
        /// Returns (bool) if path is set.
        /// Returns (string) with tmp-path to pdf if path is null
        /// </summary>
        /// <param name="path">Set Path to Target PDF.</param>
        /// <param name="sheet">Select sheet to export or 0 for All.</param>
        /// <param name="format">Select file type, can be xlsx only.</param>
        public object ExportOrSave(string? path = null, int sheet = 0, string format = "xlsx")
        {
            format = "xlsx";
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
                    if (format == "xlsx")
                        excelPackage.SaveAs(path);
                }
                else if (sheet > 0 && sheet <= excelPackage.Workbook.Worksheets.Count)
                {
                    if (format == "xlsx")
                    {
                        ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[sheet - 1];
                        using(ExcelPackage exp = new ExcelPackage())
                        {
                            exp.Workbook.Worksheets.Add(excelWorksheet.Name, excelWorksheet);
                            exp.SaveAs(path);
                        }
                        excelWorksheet.Dispose();
                    }
                }
                else return (returnPath == false ? false : "");

                return (returnPath == false ? true : path);
            }
            catch (Exception) { }

            return (returnPath == false ? false : "");
        }

        public void Dispose()
        {
            PathToSource = null;
            try { excelPackage.Dispose(); } catch (Exception) { }
            try { GC.Collect(); } catch (Exception) { }
            try { GC.WaitForPendingFinalizers(); } catch (Exception) { }
        }
    }
}
