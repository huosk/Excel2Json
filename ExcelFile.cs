using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;
using NPOI.SS.Formula.Functions;

namespace ExcelToJson
{
    public enum ColumnType
    {
        Unknown,
        Int,
        String,
        Float,
        Boolean,
    }

    public struct ColumnInfo
    {
        public int index;
        public ColumnType type;
        public bool isArray;
        public char arraySplit;
        public string name;
    }

    public class ExcelFile
    {
        public bool SkipEmptyRow { get; set; }

        private string file;
        private ColumnInfo[] columnInfos;
        private List<object[]> datas;

        public const string TYPE_STRING = "string";
        public const string TYPE_INT = "int";
        public const string TYPE_FLOAT = "float";
        public const string TYPE_BOOLEAN = "boolean";

        public ExcelFile(string file)
        {
            this.file = file;
        }

        public void ParseFile()
        {
            using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);

                // 读取第一个单元格，该单元格中保存一个数字，表示 excel 标题行号
                ICell cell = sheet.GetRow(0).GetCell(0);
                int titleRowIndex = (int)cell.NumericCellValue;

                // 因为 Excel 中指定的行号是从 1 开始的，而 GetRow 是从 0 开始的
                // 所以其实 titleRowIndex 实际指向的是 类型行
                IRow row = sheet.GetRow(titleRowIndex);
                if (row != null && row.Cells != null)
                {
                    columnInfos = new ColumnInfo[row.Cells.Count];
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        var ce = row.GetCell(i);
                        if (ce == null)
                            continue;

                        columnInfos[i].index = ce.ColumnIndex;

                        string typeStr = ce.StringCellValue.Trim();

                        var match = System.Text.RegularExpressions.Regex.Match(typeStr, @"^(\w+)\[(.)\]$");
                        if (match.Success)
                        {
                            typeStr = match.Groups[1].Value;
                            columnInfos[i].isArray = true;
                            columnInfos[i].arraySplit = match.Groups[2].Value[0];
                        }

                        if (typeStr.Equals(TYPE_STRING)) columnInfos[i].type = ColumnType.String;
                        else if (typeStr.Equals(TYPE_INT)) columnInfos[i].type = ColumnType.Int;
                        else if (typeStr.Equals(TYPE_FLOAT)) columnInfos[i].type = ColumnType.Float;
                        else if (typeStr.Equals(TYPE_BOOLEAN)) columnInfos[i].type = ColumnType.Boolean;
                        else columnInfos[i].type = ColumnType.Unknown;
                    }
                }

                // 列数以类型行为准
                int columnCount = columnInfos.Length;

                // 读取列的属性名称
                int propertyNameRowIndex = titleRowIndex + 1;
                row = sheet.GetRow(propertyNameRowIndex);
                for (int i = 0; i < columnCount; i++)
                {
                    var colInfo = columnInfos[i];
                    var ce = row.GetCell(colInfo.index);
                    if (ce == null)
                    {
                        Console.Error.WriteLine("{0} 文件的第{1}列，未设置属性名称", file, colInfo.index);
                        continue;
                    }

                    columnInfos[i].name = ce.StringCellValue;
                }

                // 开始读取数据
                datas = new List<object[]>();

                for (int rowNum = titleRowIndex + 2; rowNum <= sheet.LastRowNum; rowNum++)
                {
                    var tempRow = sheet.GetRow(rowNum);
                    if (tempRow == null || tempRow.Cells == null || tempRow.Cells.Count == 0)
                        continue;

                    if (SkipEmptyRow)
                    {
                        bool isEmptyRow = IsEmptyRow(tempRow, columnCount);
                        if (isEmptyRow)
                        {
                            continue;
                        }
                    }

                    object[] rowData = new object[columnCount];

                    for (int i = 0; i < columnCount; i++)
                    {
                        var colInfo = columnInfos[i];
                        bool isArr = columnInfos[i].isArray;
                        char split = columnInfos[i].arraySplit;

                        ICell ce = tempRow.GetCell(colInfo.index);

                        if (ce == null || (isArr && string.IsNullOrEmpty(ce.StringCellValue)))
                        {
                            rowData[i] = GetDefaultValue(columnInfos[i]);
                        }
                        else
                        {
                            if (isArr)
                                rowData[i] = ce.StringCellValue.Split(split).Select((v) => ConvertCell(columnInfos[i].type, v)).ToArray();
                            else
                                rowData[i] = GetCellValue(columnInfos[i].type, ce);
                        }
                    }

                    datas.Add(rowData);
                }
            }
        }

        // 检测是否为空行
        private bool IsEmptyRow(IRow row, int columnCount)
        {
            if (row == null || row.Cells == null || row.Cells.Count == 0)
                return true;

            for (int i = 0; i < columnCount; i++)
            {
                var cell = row.GetCell(i);
                if (cell == null)
                    continue;

                if (cell.CellType == CellType.Blank)
                    continue;
                else if (cell.CellType == CellType.String && !string.IsNullOrEmpty(cell.StringCellValue))
                    return false;
            }

            return true;
        }

        private object GetCellValue(ColumnType type, ICell ce)
        {
            if (type == ColumnType.Int) return (int)ce.NumericCellValue;
            else if (type == ColumnType.Float) return (float)ce.NumericCellValue;
            else if (type == ColumnType.String) return ce.StringCellValue;
            else if (type == ColumnType.Boolean) return ce.BooleanCellValue;
            else return ce.StringCellValue;
        }

        private object ConvertCell(ColumnType type, string s)
        {
            if (type == ColumnType.Int) return Convert.ToInt32(s);
            else if (type == ColumnType.Float) return Convert.ToSingle(s);
            else if (type == ColumnType.String) return s;
            else if (type == ColumnType.Boolean) return Convert.ToBoolean(s);
            else return s;
        }

        private object GetDefaultValue(ColumnInfo inf)
        {
            if (inf.isArray)
                return null;

            switch (inf.type)
            {
                case ColumnType.Int: return 0;
                case ColumnType.Float: return 0f;
                case ColumnType.String: return string.Empty;
                case ColumnType.Boolean: return false;
                default:
                    return null;
            }
        }

        public List<KeyValuePair<string, object>[]> GetRows()
        {
            return datas.Select((row) =>
            {
                return row.Select((cell, index) =>
                 {
                     return new KeyValuePair<string, object>(columnInfos[index].name, cell);
                 }).ToArray();
            }).ToList();
        }
    }
}
