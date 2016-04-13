using System;
using System.Collections.Generic;
using System.Data;
using OfficeOpenXml;
using System.IO;
using ExcelExtentions.Argument;
using System.Reflection;

namespace ExcelExtentions {
    internal class ExcelService {
        public static void Export(DataTable dt, ExcelArgument arg) {
            using (ExcelPackage ep = new ExcelPackage()) {
                ep.DoAdjustDrawings = false;
                ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(arg.SheetArguments[0].SheetName);
                createSheet(worksheet, dt, arg.SheetArguments[0]);
                createFile(ep, arg.OutPutPath);
            }
        }
        public static void Export(DataSet ds, ExcelArgument arg) {
            using (ExcelPackage ep = new ExcelPackage()) {
                ep.DoAdjustDrawings = false;
                for (int i = 0; i < ds.Tables.Count; i++) {
                    ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(arg.SheetArguments[i].SheetName);
                    createSheet(worksheet, ds.Tables[i], arg.SheetArguments[i]);
                }
                createFile(ep, arg.OutPutPath);
            }
        }
        public static void Export(dynamic list, ExcelArgument arg) {
            using (ExcelPackage ep = new ExcelPackage()) {
                ep.DoAdjustDrawings = false;
                ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(arg.SheetArguments[0].SheetName);
                createSheet(worksheet, list, arg.SheetArguments[0]);
                createFile(ep, arg.OutPutPath);
            }
        }
        public static void Export(List<dynamic> list, ExcelArgument arg) {
            using (ExcelPackage ep = new ExcelPackage()) {
                ep.DoAdjustDrawings = false;
                for (int i = 0; i < list.Count; i++) {
                    ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(arg.SheetArguments[i].SheetName);
                    createSheet(worksheet, list[i], arg.SheetArguments[i]);
                }
                createFile(ep, arg.OutPutPath);
            }
        }
        private static void createSheet(ExcelWorksheet worksheet, DataTable dt, SheetArgument arg) {
            int magicToken = 0;
            int colCount = dt.Columns.Count;
            int rowCount = dt.Rows.Count;
            int rowNumber = 1;
            worksheet.Row(rowNumber).Height = arg.TitleHeight;
            for (int i = 0; i < colCount; ++i) {
                worksheet.Column(i + 1).Width = arg.ColumnArguments[i].Width;
                worksheet.Cells[rowNumber, i + 1].Value = dt.Columns[i].ColumnName;
                setColStyle(worksheet.Column(i + 1), arg.ColumnArguments[i].ColumnValueType);
            }
            for (int i = 0; i < rowCount; ++i) {
                rowNumber++;
                DataRow dr = dt.Rows[i];
                worksheet.Row(rowNumber).Height = arg.RowHeight;
                for (int ii = 0; ii < colCount; ++ii) {
                    setCellValue(worksheet, rowNumber, ii + 1, dr[ii], arg.ColumnArguments[ii].ColumnValueType, ref magicToken);
                }
            }
        }
        private static void createSheet(ExcelWorksheet worksheet, dynamic list, SheetArgument arg) {
            int magicToken = 0;
            Type listType = list.GetType();
            MethodInfo m_Count = listType.GetMethod("Count");
            Type type = null;
            foreach (var l in list) {
                type = l.GetType();
                break;
            }
            var properties = type.GetProperties();
            int colCount = properties.Length;
            int rowNumber = 1;
            worksheet.Row(rowNumber).Height = arg.TitleHeight;
            for (int i = 0; i < colCount; ++i) {
                worksheet.Column(i + 1).Width = arg.ColumnArguments[i].Width;
                worksheet.Cells[rowNumber, i + 1].Value = properties[i].Name;
                setColStyle(worksheet.Column(i + 1), arg.ColumnArguments[i].ColumnValueType);
            }
            foreach (var l in list) {
                rowNumber++;
                var row = l;
                worksheet.Row(rowNumber).Height = arg.RowHeight;
                for (int ii = 0; ii < colCount; ++ii) {
                    setCellValue(worksheet, rowNumber, ii + 1, properties[ii].GetValue(row), arg.ColumnArguments[ii].ColumnValueType, ref magicToken);
                }
            }
        }
        private static void setColStyle(ExcelColumn col, ColumnValueType type) {
            switch (type) {
                //case ColumnValueType.String:
                //    break;
                //case ColumnValueType.Int:
                //    break;
                //case ColumnValueType.Double:
                //    break;
                //case ColumnValueType.Picture:
                //    break;
                //case ColumnValueType.IntNull:
                //    break;
                //case ColumnValueType.DoubleNull:
                //    break;
                case ColumnValueType.DateTime:
                    col.Style.Numberformat.Format = "yyyy/m/d  hh:mm:ss";
                    break;
                case ColumnValueType.Date:
                    col.Style.Numberformat.Format = "yyyy/m/d";
                    break;
                case ColumnValueType.Time:
                    col.Style.Numberformat.Format = "hh:mm:ss";
                    break;
                //case ColumnValueType.Currency:
                //    col.Style.Numberformat.Format = string.Format("{0}#,##0.00",code);
                //    break;
                default: break;
            }
        }
        private static void setCellValue(ExcelWorksheet workSheet, int row, int col, object value, ColumnValueType type, ref int magicToken) {
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
                    string picPath = value.ToString();
                    if (File.Exists(picPath)) {
                        ExcelPictureProcessor.SetPictureToCell(picPath, workSheet, row, col, magicToken);
                        magicToken++;
                    }
                    else {
                        cell.Value = "No Pic";
                    }
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
                    var t = splitCurrencyValue(value.ToString());
                    cell.Style.Numberformat.Format = string.Format("[${0}] #0.00", t.Item1);
                    cell.Value = t.Item2;
                    break;
                default: break;
            }
        }
        private static void createFile(ExcelPackage ep, string fileName) {
            FileInfo newFile = new FileInfo(fileName);
            if (newFile.Exists) {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(fileName);
            }
            ep.SaveAs(newFile);
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
    }
}
