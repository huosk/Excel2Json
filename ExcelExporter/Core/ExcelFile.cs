using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using NPOI.SS.Formula.Functions;
using NPOI.HSSF.UserModel;

namespace ExcelExporter
{
    public struct ColumnInfo
    {
        public int number;
        public string name;
    }

    public class ExcelFile
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetNumber"></param>
        /// <param name="titleRowNum"></param>
        /// <param name="dataRowNum"></param>
        /// <returns></returns>
        public static List<Dictionary<string, object>> ParseFile(string file, int sheetNumber, int titleRowNum, int dataRowNum)
        {
            List<Dictionary<string, object>> datas = new List<Dictionary<string, object>>();

            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = WorkbookFactory.Create(fs, true);
                ISheet sheet = workbook.GetSheetAt(sheetNumber);
                IFormulaEvaluator evaluator = WorkbookFactory.CreateFormulaEvaluator(workbook);

                datas = ParseSheet(sheet, titleRowNum, dataRowNum, evaluator);

                workbook.Close();
            }

            return datas;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetName"></param>
        /// <param name="titleRowNum"></param>
        /// <param name="dataRowNum"></param>
        /// <returns></returns>
        public static List<Dictionary<string, object>> ParseFile(string file, string sheetName, int titleRowNum, int dataRowNum)
        {
            List<Dictionary<string, object>> datas = new List<Dictionary<string, object>>();

            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = WorkbookFactory.Create(fs, true);
                ISheet sheet = workbook.GetSheet(sheetName);
                IFormulaEvaluator evaluator = WorkbookFactory.CreateFormulaEvaluator(workbook);

                datas = ParseSheet(sheet, titleRowNum, dataRowNum, evaluator);

                workbook.Close();
            }

            return datas;
        }

        private static List<Dictionary<string, object>> ParseSheet(ISheet sheet, int titleRowNum, int dataRowNum, IFormulaEvaluator evaluator)
        {
            List<Dictionary<string, object>> datas = new List<Dictionary<string, object>>();

            var columnInfos = ParseTitle(sheet, titleRowNum);

            // 开始读取数据
            int realDataRowNum = Math.Max(dataRowNum, sheet.FirstRowNum);
            for (int rowNum = realDataRowNum; rowNum <= sheet.LastRowNum; rowNum++)
            {
                var tempRow = sheet.GetRow(rowNum);
                if (tempRow == null || tempRow.Cells == null || tempRow.Cells.Count == 0)
                    continue;

                if (IsEmptyRow(tempRow))
                    continue;

                Dictionary<string, object> rowMap = ParseRow(tempRow, columnInfos, evaluator);
                if (rowMap != null) datas.Add(rowMap);
            }

            return datas;
        }

        private static ColumnInfo[] ParseTitle(ISheet sheet, int titleRowNum)
        {
            IRow titleRow = sheet.GetRow(titleRowNum);
            var colInfs = new List<ColumnInfo>();
            if (titleRow != null)
            {
                for (int i = 0; i < titleRow.Cells.Count; i++)
                {
                    var ce = titleRow.Cells[i];
                    if (ce == null)
                        continue;

                    colInfs.Add(new ColumnInfo()
                    {
                        number = ce.ColumnIndex,
                        name = ce.StringCellValue
                    });
                }
            }
            return colInfs.ToArray();
        }

        private static Dictionary<string, object> ParseRow(IRow row, ColumnInfo[] titles, IFormulaEvaluator evaluator)
        {
            Dictionary<string, object> rowMap = new Dictionary<string, object>();

            for (int i = 0; i < titles.Length; i++)
            {
                ICell cell = row.GetCell(titles[i].number);
                object cellVal = null;
                try
                {
                    ICell ce = cell;
                    if (ce != null)
                    {
                        if (ce.CellType == CellType.Formula)
                        {
                            ce = evaluator.EvaluateInCell(ce);
                        }

                        switch (ce.CellType)
                        {
                            case CellType.Numeric:
                            case CellType.Formula:
                                cellVal = ce.NumericCellValue;
                                break;
                            case CellType.Unknown:
                            case CellType.String:
                                cellVal = ce.StringCellValue;
                                break;
                            case CellType.Boolean:
                                cellVal = ce.BooleanCellValue;
                                break;
                            case CellType.Error:
                                cellVal = ce.ErrorCellValue;
                                break;
                        }
                    }

                    rowMap.Add(titles[i].name, cellVal);
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine("({0}):{1}", cell.Address.ToString(), e);
                    return null;
                }
            }

            return rowMap;
        }

        // 检测是否为空行
        private static bool IsEmptyRow(IRow row)
        {
            if (row == null || row.Cells == null || row.Cells.Count == 0)
                return true;

            int minColNum = row.FirstCellNum;
            int maxColNum = row.LastCellNum;
            for (int i = minColNum; i < maxColNum; i++)
            {
                var cell = row.GetCell(i);
                if (cell == null)
                    continue;

                if (cell.CellType == CellType.Blank)
                    continue;
                else if (cell.CellType == CellType.String && !string.IsNullOrEmpty(cell.StringCellValue))
                    return false;
                else
                    return false;
            }

            return true;
        }
    }
}