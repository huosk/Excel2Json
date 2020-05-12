using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Linq;

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

    public class ExcelFile
    {
        private string file;

        private ColumnType[] columnTypes;
        private List<object[]> datas;
        private string[] columns;

        public const string TYPE_STRING = "string";
        public const string TYPE_INT = "int";
        public const string TYPE_FLOAT = "float";
        public const string TYPE_BOOLEAN = "boolean";

        public ExcelFile(string file)
        {
            this.file = file;

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
                    columnTypes = new ColumnType[row.Cells.Count];
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        string typeStr = row.GetCell(i).StringCellValue;
                        if (typeStr.Equals(TYPE_STRING)) columnTypes[i] = ColumnType.String;
                        else if (typeStr.Equals(TYPE_INT)) columnTypes[i] = ColumnType.Int;
                        else if (typeStr.Equals(TYPE_FLOAT)) columnTypes[i] = ColumnType.Float;
                        else if (typeStr.Equals(TYPE_BOOLEAN)) columnTypes[i] = ColumnType.Boolean;
                        else columnTypes[i] = ColumnType.Unknown;
                    }
                }

                // 列数以类型行为准
                int columnCount = columnTypes.Length;

                // 读取列的属性名称
                int propertyNameRowIndex = titleRowIndex + 1;
                columns = new string[columnCount];
                row = sheet.GetRow(propertyNameRowIndex);
                for (int i = 0; i < columnCount; i++)
                {
                    columns[i] = row.GetCell(i).StringCellValue;
                }

                // 开始读取数据
                int dataIndex = titleRowIndex + 2;
                row = sheet.GetRow(dataIndex);
                datas = new List<object[]>();

                while (row != null && row.Cells != null && row.Cells.Count > 0)
                {
                    object[] rowData = new object[columnCount];

                    for (int i = 0; i < columnCount; i++)
                    {
                        if (i < row.Cells.Count)
                        {
                            var ce = row.GetCell(i);
                            switch (columnTypes[i])
                            {
                                case ColumnType.Int:
                                    rowData[i] = (int)ce.NumericCellValue;
                                    break;
                                case ColumnType.String:
                                    rowData[i] = ce.StringCellValue;
                                    break;
                                case ColumnType.Float:
                                    rowData[i] = (float)ce.NumericCellValue;
                                    break;
                                case ColumnType.Boolean:
                                    rowData[i] = ce.BooleanCellValue;
                                    break;
                                default:
                                    rowData[i] = ce.StringCellValue;
                                    break;
                            }
                        }
                        else
                        {
                            rowData[i] = GetDefaultValue(columnTypes[i]);
                        }
                    }

                    datas.Add(rowData);

                    row = sheet.GetRow(++dataIndex);
                }
            }
        }

        private object GetDefaultValue(ColumnType type)
        {
            switch (type)
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
                     return new KeyValuePair<string, object>(columns[index], cell);
                 }).ToArray();
            }).ToList();
        }
    }
}
