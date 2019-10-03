using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOICSVToExcel
{
    using System.Diagnostics;
    using System.IO;

    public class NPOICSVToExcel
    {
        public static string Convert(string csvpath, string excelpath, string worksheetName, int totalCols = 0)
        {
            IEnumerable<string[]> csvLines;
            try
            {
                csvLines = System.IO.File.ReadAllLines(csvpath, Encoding.UTF8).Select(a => a.Split(';'));

                if (csvLines == null || csvLines.Count() == 0)
                {
                    return ("Empty file");
                }

                int rowCount = 0;
                int colCount = 0;


                IWorkbook workbook = new XSSFWorkbook();
                ISheet worksheet = workbook.CreateSheet(worksheetName);

                foreach (var line in csvLines)
                {
                    IRow row = worksheet.CreateRow(rowCount);

                    colCount = 0;
                    foreach (var col in line)
                    {
                        row.CreateCell(colCount).SetCellValue(col);
                        //row.CreateCell(colCount).SetCellValue(TypeConverter.TryConvert(col));


                        //DateTime dateTimeValue = DateTime.MinValue;
                        ////decimal decimalValue = 0;
                        //double doubleValue = 0;
                        //float floatValue = 0;
                        //long longValue = 0;
                        //int intValue = 0;
                        //bool boolValue = false;

                        //if (string.IsNullOrEmpty(col))
                        //{
                        //    row.CreateCell(colCount).SetCellValue(string.Empty);
                        //}
                        ////else
                        ////if (decimal.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture.NumberFormat, out decimalValue))
                        ////{
                        ////    row.CreateCell(colCount).SetCellValue(decimalValue);
                        ////}
                        //else
                        //if (double.TryParse(col, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture.NumberFormat, out doubleValue))
                        //{
                        //    row.CreateCell(colCount).SetCellValue(doubleValue);
                        //}
                        //else
                        //if (float.TryParse(col, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture.NumberFormat, out floatValue))
                        //{
                        //    row.CreateCell(colCount).SetCellValue(floatValue);
                        //}
                        //else                        
                        //if (long.TryParse(col, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture.NumberFormat, out longValue))
                        //{
                        //    row.CreateCell(colCount).SetCellValue(longValue);
                        //}
                        //else                        
                        //if (int.TryParse(col, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture.NumberFormat, out intValue))
                        //{
                        //    row.CreateCell(colCount).SetCellValue(intValue);
                        //}
                        //else
                        //if (bool.TryParse(col, out boolValue))
                        //{
                        //    row.CreateCell(colCount).SetCellValue(boolValue);
                        //}
                        //else
                        //if (DateTime.TryParse(col, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTimeValue))
                        //{
                        //    row.CreateCell(colCount).SetCellValue(dateTimeValue);
                        //}
                        //else
                        //{
                        //    row.CreateCell(colCount).SetCellValue(col);
                        //}

                        colCount++;
                    }
                    rowCount++;
                }

                using (FileStream fileWriter = File.Create(excelpath))
                {
                    workbook.Write(fileWriter);
                    fileWriter.Close();
                }

                worksheet = null;
                workbook = null;
                
                return "OK";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
    }
}
