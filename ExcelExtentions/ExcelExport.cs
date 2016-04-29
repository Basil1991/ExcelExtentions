using System.Collections.Generic;
using System.Data;
using OfficeOpenXml;
using ExcelExtentions.Argument;

namespace ExcelExtentions {
    internal class ExcelExport {
        public static void Export(DataTable dt, ExcelArgument arg) {
            using (ExcelPackage ep = new ExcelPackage()) {
                ep.DoAdjustDrawings = false;
                ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(arg.SheetArguments[0].SheetName);
                ExcelSheetCreator.CreateSheet(worksheet, dt, arg.SheetArguments[0]);
                ExcelCommonService.CreateFile(ep, arg.OutPutPath);
            }
        }
        public static void Export(DataSet ds, ExcelArgument arg) {
            using (ExcelPackage ep = new ExcelPackage()) {
                ep.DoAdjustDrawings = false;
                for (int i = 0; i < ds.Tables.Count; i++) {
                    ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(arg.SheetArguments[i].SheetName);
                    ExcelSheetCreator.CreateSheet(worksheet, ds.Tables[i], arg.SheetArguments[i]);
                }
                ExcelCommonService.CreateFile(ep, arg.OutPutPath);
            }
        }
        public static void Export(dynamic list, ExcelArgument arg) {
            using (ExcelPackage ep = new ExcelPackage()) {
                ep.DoAdjustDrawings = false;
                ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(arg.SheetArguments[0].SheetName);
                ExcelSheetCreator.CreateSheet(worksheet, list, arg.SheetArguments[0]);
                ExcelCommonService.CreateFile(ep, arg.OutPutPath);
            }
        }
        public static void Export(List<dynamic> list, ExcelArgument arg) {
            using (ExcelPackage ep = new ExcelPackage()) {
                ep.DoAdjustDrawings = false;
                for (int i = 0; i < list.Count; i++) {
                    ExcelWorksheet worksheet = ep.Workbook.Worksheets.Add(arg.SheetArguments[i].SheetName);
                    ExcelSheetCreator.CreateSheet(worksheet, list[i], arg.SheetArguments[i]);
                }
                ExcelCommonService.CreateFile(ep, arg.OutPutPath);
            }
        }
    }
}
