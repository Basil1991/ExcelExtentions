using ExcelExtentions.Core.Argument;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;

namespace ExcelExtentions.Core {
    internal class ExcelSheetCreator {
        public static void CreateSheet(ExcelWorksheet worksheet, DataTable dt, SheetArgument arg) {
            int colCount = dt.Columns.Count;
            int rowCount = dt.Rows.Count;
            int rowNumber = 1;
            worksheet.Row(rowNumber).Height = arg.TitleHeight;
            for (int i = 0; i < colCount; ++i) {
                worksheet.Cells[rowNumber, i + 1].Value = dt.Columns[i].ColumnName;
                ExcelStyleProcessor.SetColStyle(worksheet, i, arg);
            }
            for (int i = 0; i < rowCount; ++i) {
                rowNumber++;
                DataRow dr = dt.Rows[i];
                worksheet.Row(rowNumber).Height = arg.RowHeight;
                for (int ii = 0; ii < colCount; ++ii) {
                    ExcelCellValueProcessor.Set(worksheet, rowNumber, ii + 1, dr[ii], arg.ColumnArguments[ii].ColumnValueType);
                }
            }
        }
        public static void CreateSheet(ExcelWorksheet worksheet, dynamic list, SheetArgument arg) {
            var properties = ExcelCommonService.PInfoBydynamics(list);
            int colCount = properties.Length;
            int rowNumber = 1;
            worksheet.Row(rowNumber).Height = arg.TitleHeight;
            for (int i = 0; i < colCount; ++i) {
                worksheet.Cells[rowNumber, i + 1].Value = properties[i].Name;
                ExcelStyleProcessor.SetColStyle(worksheet, i, arg);
            }
            foreach (var l in list) {
                rowNumber++;
                var row = l;
                worksheet.Row(rowNumber).Height = arg.RowHeight;
                for (int ii = 0; ii < colCount; ++ii) {
                    ExcelCellValueProcessor.Set(worksheet, rowNumber, ii + 1, properties[ii].GetValue(row), arg.ColumnArguments[ii].ColumnValueType);
                }
            }
        }
        public static void CreateSheet(ExcelWorksheet worksheet, DataTable dt) {
            int colCount = dt.Columns.Count;
            int rowCount = dt.Rows.Count;
            int rowNumber = 1;
            for (int i = 0; i < rowCount; ++i) {
                rowNumber++;
                DataRow dr = dt.Rows[i];
                for (int ii = 0; ii < colCount; ++ii) {
                    ExcelCellValueProcessor.Set(worksheet, rowNumber, ii + 1, dr[ii]);
                }
            }
        }
        public static void CreateSheet(ExcelWorksheet worksheet, DataTable dt, ICollection<int> pcs) {
            int magicToken = 0;
            int colCount = dt.Columns.Count;
            int rowCount = dt.Rows.Count;
            int rowNumber = 1;
            for (int i = 0; i < rowCount; ++i) {
                rowNumber++;
                DataRow dr = dt.Rows[i];
                for (int ii = 0; ii < colCount; ++ii) {
                    ExcelCellValueProcessor.Set(worksheet, rowNumber, ii + 1, dr[ii], pcs, ref magicToken);
                }
            }
        }
        public static void CreateSheet(ExcelWorksheet worksheet, dynamic list) {
            var properties = ExcelCommonService.PInfoBydynamics(list);
            int colCount = properties.Length;
            int rowNumber = 1;
            foreach (var l in list) {
                rowNumber++;
                var row = l;
                for (int ii = 0; ii < colCount; ++ii) {
                    ExcelCellValueProcessor.Set(worksheet, rowNumber, ii + 1, properties[ii].GetValue(row));
                }
            }
        }
        public static void CreateSheet(ExcelWorksheet worksheet, dynamic list, ICollection<int> pcs) {
            int magicToken = 0;
            var properties = ExcelCommonService.PInfoBydynamics(list);
            int colCount = properties.Length;
            int rowNumber = 1;
            foreach (var l in list) {
                rowNumber++;
                var row = l;
                for (int ii = 0; ii < colCount; ++ii) {
                    ExcelCellValueProcessor.Set(worksheet, rowNumber, ii + 1, properties[ii].GetValue(row), pcs, ref magicToken);
                }
            }
        }
    }
}
