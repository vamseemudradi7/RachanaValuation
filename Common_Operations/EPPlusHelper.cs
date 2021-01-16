using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DataTable = System.Data.DataTable;

namespace BhavCopy
{
    public class EpPlusHelper
    {
        public static T ReadExcel<T>(string path, string sheetName)
        {
            DataTable dtTable = new DataTable();
            List<string> rowList = new List<string>();
            ISheet sheet;
            using (var stream = new FileStream(path, FileMode.Open))
            {
                stream.Position = 0;
                HSSFWorkbook xssWorkbook = new HSSFWorkbook(stream);
                sheet = xssWorkbook.GetSheet(sheetName);
                IRow headerRow = sheetName == "ALF suspended" ? sheet.GetRow(1) : sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum;
                for (int j = 0; j < cellCount; j++)
                {
                    ICell cell = headerRow.GetCell(j);
                    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                    {
                        dtTable.Columns.Add(cell.ToString());
                    }
                }
                for (int i = sheetName == "ALF suspended" ? (sheet.FirstRowNum + 2) : (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                            {
                                rowList.Add(row.GetCell(j).ToString());
                            }
                        }
                    }
                    if (rowList.Count > 0)
                        dtTable.Rows.Add(rowList.ToArray());
                    rowList.Clear();
                }
            }
            var generatedType = JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(dtTable));
            return (T)Convert.ChangeType(generatedType, typeof(T));
        }

        public T ReadFromExcel<T>(string path, string workSheetName, bool hasHeader = true)
        {
            using var excelPack = new ExcelPackage();
            //Load excel stream
            using (var stream = File.OpenRead(path))
            {
                excelPack.Load(stream);
            }

            //Lets Deal with first worksheet.(You may iterate here if dealing with multiple sheets)
            var ws = excelPack.Workbook.Worksheets.First(x => x.Name == workSheetName);

            //Get all details as DataTable -because Datatable make life easy :)
            DataTable excelasTable = new DataTable();
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                //Get colummn details
                if (!string.IsNullOrEmpty(firstRowCell.Text))
                {
                    string firstColumn = string.Format("Column {0}", firstRowCell.Start.Column);
                    excelasTable.Columns.Add(hasHeader ? firstRowCell.Text : firstColumn);
                }
            }
            var startRow = hasHeader ? 2 : 1;
            //Get row details
            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, excelasTable.Columns.Count];
                DataRow row = excelasTable.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }
            //Get everything as generics and let end user decides on casting to required type
            var generatedType = JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(excelasTable));
            return (T)Convert.ChangeType(generatedType, typeof(T));
        }
    }
}
