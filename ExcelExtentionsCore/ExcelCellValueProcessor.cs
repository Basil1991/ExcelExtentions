using ExcelExtentions.Core.Argument;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelExtentions.Core {
    public class ExcelCellValueProcessor {
        public static void Set(ExcelWorksheet workSheet, int row, int col, object value, ColumnValueType type) {
            var cell = workSheet.Cells[row, col];
            switch (type) {
                case ColumnValueType.String:
                    cell.Value = value != DBNull.Value && value != null ? Convert.ToString(value) : "";
                    break;
                case ColumnValueType.Int:
                    cell.Value = value != DBNull.Value && value != null ? Convert.ToInt32(value) : 0;
                    break;
                case ColumnValueType.Double:
                    cell.Value = value != DBNull.Value && value != null ? Convert.ToDouble(value) : 0.00;
                    break;
                case ColumnValueType.DateTime:
                    if (value != DBNull.Value && value != null) {
                        cell.Value = Convert.ToDateTime(value);
                    }
                    else {
                        cell.Value = "";
                    }
                    break;
                case ColumnValueType.Picture:
                    setPicture(workSheet, cell, row, col, value);
                    break;
                case ColumnValueType.IntNull:
                    if (value != DBNull.Value && value != null) {
                        cell.Value = Convert.ToInt32(value);
                    }
                    else {
                        cell.Value = "";
                    }
                    break;
                case ColumnValueType.DoubleNull:
                    if (value != DBNull.Value && value != null) {
                        cell.Value = Convert.ToDouble(value);
                    }
                    else {
                        cell.Value = "";
                    }
                    break;
                case ColumnValueType.Currency:
                    var str = value == null || value == DBNull.Value ? "0.00" : value.ToString();
                    var t = splitCurrencyValue(str);
                    cell.Style.Numberformat.Format = string.Format("[${0}] #0.00", t.Item1);
                    cell.Value = t.Item2;
                    break;
                default: break;
            }
        }
        public static void Set(ExcelWorksheet workSheet, int row, int col, object value, ICollection<int> picCol, ref int magicToken) {
            var cell = workSheet.Cells[row, col];
            if (picCol.Contains(col)) {
                setPicture(workSheet, cell, row, col, value);
            }
            else {
                cell.Value = value;
            }
        }
        public static void Set(ExcelWorksheet workSheet, int row, int col, object value) {
            var cell = workSheet.Cells[row, col];
            cell.Value = value;
        }
        private static Tuple<string, decimal> splitCurrencyValue(string value) {
            string code = "";
            int firstIndexOfDigit = 0;
            for (int i = 0; i < value.Length; i++) {
                if (char.IsLetter(value[i])) {
                    code += value[i];
                }
                else if (char.IsDigit(value[i])) {
                    firstIndexOfDigit = i;
                    break;
                }
            }
            if (firstIndexOfDigit == 0) {
                return new Tuple<string, decimal>(code, (decimal)0.00);
            }
            else {
                return new Tuple<string, decimal>(code, Decimal.Parse(value.Substring(firstIndexOfDigit)));
            }
        }
        private static void setPicture(ExcelWorksheet workSheet, ExcelRange cell, int row, int col, object value) {
            string picPath = value == null || value == DBNull.Value ? "" : value.ToString();
            if (Uri.IsWellFormedUriString(picPath, UriKind.Absolute)) {
                try {
                    System.Net.WebRequest webreq = System.Net.WebRequest.Create(picPath);
                    System.Net.WebResponse webres = webreq.GetResponse();
                    var stream = webres.GetResponseStream();
                    ExcelPictureProcessor.SetPictureToCell(stream, workSheet, row, col);
                }
                catch {
                    cell.Value = "No Pic";
                }
            }
            else if (File.Exists(picPath)) {
                ExcelPictureProcessor.SetPictureToCell(picPath, workSheet, row, col);
            }
            else {
                cell.Value = "No Pic";
            }
        }
    }
}
