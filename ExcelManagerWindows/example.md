# Example 1: Open xlsx file, change Cell content and export as pdf or xlsx
using (Exel excel = new Exel("Path to xlsx or byte[] from xlsx file"))
{
  excel.UpdateCell("A", "1", "New Cell Text", 1);
  Console.Write(excel.ExportOrSave(null, 0, "xlsx"));
  Console.Write(excel.ExportOrSave(null, 0, "pdf"));
}
