using ExcelExtentions.Argument;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtentions {
    internal class ExcelStyleProcessor {
        public static void SetColStyle(ExcelWorksheet worksheet, int colNum, SheetArgument arg) {
            worksheet.Column(colNum + 1).Width = arg.ColumnArguments[colNum].Width;
            setStyle(worksheet.Column(colNum + 1), arg.ColumnArguments[colNum].ColumnValueType);
            setClass(worksheet.Column(colNum + 1), arg.ClassType);
        }
        private static void setStyle(ExcelColumn col, ColumnValueType type) {
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
        private static void setClass(ExcelColumn col, ClassType type) {
            switch (type) {
                case ClassType.Default: break;
                case ClassType.AllCenter:
                    col.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    col.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    break;
                default: break;
            }
        }
    }
}
