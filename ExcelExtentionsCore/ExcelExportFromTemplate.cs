using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;

namespace ExcelExtentions.Core {
    internal class ExcelExportFromTemplate {
        public static void ExportFromTemp(DataTable dt, string tempPath, string outPutPath) {
            var fs = ExcelCommonService.GetFileSteam(tempPath);
            using (ExcelPackage ep = new ExcelPackage(fs)) {
                var pcs = ExcelCommonService.PicCols(dt);
                ep.DoAdjustDrawings = false;
                if (pcs.Count == 0)
                    ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[1], dt);
                else
                    ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[1], dt, pcs);
                ExcelCommonService.CreateFile(ep, outPutPath);
            }
        }
        public static void ExportFromTemp(DataSet ds, string tempPath, string outPutPath) {
            var fs = ExcelCommonService.GetFileSteam(tempPath);
            using (ExcelPackage ep = new ExcelPackage(fs)) {
                ep.DoAdjustDrawings = false;
                for (int i = 0; i < ds.Tables.Count; i++) {
                    var dt = ds.Tables[i];
                    var pcs = ExcelCommonService.PicCols(dt);
                    if (pcs.Count == 0)
                        ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[i + 1], dt);
                    else
                        ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[i + 1], dt, pcs);
                }
                ExcelCommonService.CreateFile(ep, outPutPath);
            }
        }
        public static void ExportFromTemp(dynamic list, string tempPath, string outPutPath) {
            var fs = ExcelCommonService.GetFileSteam(tempPath);
            using (ExcelPackage ep = new ExcelPackage(fs)) {
                var pcs = ExcelCommonService.PicCols(list);
                ep.DoAdjustDrawings = false;
                if (pcs.Count == 0)
                    ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[1], list);
                else
                    ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[1], list, pcs);
                ExcelCommonService.CreateFile(ep, outPutPath);
            }
        }
        public static void ExportFromTemp(List<dynamic> list, string tempPath, string outPutPath) {
            var fs = ExcelCommonService.GetFileSteam(tempPath);
            using (ExcelPackage ep = new ExcelPackage(fs)) {
                ep.DoAdjustDrawings = false;
                for (int i = 0; i < list.Count; i++) {
                    dynamic d = list[i];
                    var pcs = ExcelCommonService.PicCols(d);
                    if (pcs.Count == 0)
                        ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[i + 1], d);
                    else
                        ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[i + 1], d, pcs);
                }

                ExcelCommonService.CreateFile(ep, outPutPath);
            }
        }
        public static void ExportFromTemp<T>(List<List<T>> list, string tempPath, string outPutPath) {
            var fs = ExcelCommonService.GetFileSteam(tempPath);
            using (ExcelPackage ep = new ExcelPackage(fs)) {
                ep.DoAdjustDrawings = false;
                for (int i = 0; i < list.Count; i++) {
                    dynamic d = list[i];
                    var pcs = ExcelCommonService.PicCols(d);
                    if (pcs.Count == 0)
                        ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[i + 1], d);
                    else
                        ExcelSheetCreator.CreateSheet(ep.Workbook.Worksheets[i + 1], d, pcs);
                }
                ExcelCommonService.CreateFile(ep, outPutPath);
            }
        }
    }
}
