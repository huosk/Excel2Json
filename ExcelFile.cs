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

    public struct ColumnInfo
    {
        public int index;
        public ColumnType type;
        public bool isArray;
        public char arraySplit;
    }

    public class ExcelFile
    {
        private string file;

        private ColumnInfo[] columnInfos;
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
                columns = new string[columnCount];
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

                    columns[i] = ce.StringCellValue;
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
                        var colInfo = columnInfos[i];
                        ICell ce = row.GetCell(colInfo.index);

                        if (ce != null)
                        {
                            bool isArr = columnInfos[i].isArray;
                            char split = columnInfos[i].arraySplit;
                            switch (columnInfos[i].type)
                            {
                                case ColumnType.Int:
                                    if (isArr)
                                        rowData[i] = ce.StringCellValue.Split(split).Select((v) =>
                                        {
                                            int num = 0;
                                            int.TryParse(v, out num);
                                            return num;
                                        }).ToArray();
                                    else
                                        rowData[i] = (int)ce.NumericCellValue;
                                    break;
                                case ColumnType.String:
                                    if (isArr)
                                        rowData[i] = ce.StringCellValue.Split(split);
                                    else
                                        rowData[i] = ce.StringCellValue;
                                    break;
                                case ColumnType.Float:
                                    if (isArr)
                                        rowData[i] = ce.StringCellValue.Split(split).Select((v) =>
                                        {
                                            float num = 0.0f;
                                            float.TryParse(v, out num);
                                            return num;
                                        }).ToArray();
                                    else
                                        rowData[i] = (float)ce.NumericCellValue;
                                    break;
                                case ColumnType.Boolean:
                                    if (isArr)
                                        rowData[i] = ce.StringCellValue.Split(split).Select((v) =>
                                        {
                                            bool b = false;
                                            bool.TryParse(v, out b);
                                            return b;
                                        }).ToArray();
                                    else
                                        rowData[i] = ce.BooleanCellValue;
                                    break;
                                default:
                                    rowData[i] = ce.StringCellValue;
                                    break;
                            }
                        }
                        else
                        {
                            rowData[i] = GetDefaultValue(columnInfos[i].type);
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
