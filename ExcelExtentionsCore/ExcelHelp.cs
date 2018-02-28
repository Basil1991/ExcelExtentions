using System.Collections.Generic;
using System.Data;
using ExcelExtentions.Core.Argument;

namespace ExcelExtentions.Core {
    public class ExcelHelp {
        public void Get(DataTable dt, ExcelArgument arg) {
            ExcelExport.Export(dt, arg);
        }
        public void Get(DataSet ds, ExcelArgument arg) {
            ExcelExport.Export(ds, arg);
        }
        public void Get(dynamic list, ExcelArgument arg) {
            ExcelExport.Export(list, arg);
        }
        public void Get(List<dynamic> lists, ExcelArgument arg) {
            ExcelExport.Export(lists, arg);
        }
        public void Get<T>(List<T> list, ExcelArgument arg) {
            ExcelExport.Export(list, arg);
        }
        public void Get<T>(List<List<T>> lists, ExcelArgument arg) {
            ExcelExport.Export<T>(lists, arg);
        }
        public void GetFromTemplate(DataTable dt, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(dt, tempPath, outPutPath);
        }
        public void GetFromTemplate(DataSet ds, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(ds, tempPath, outPutPath);
        }
        public void GetFromTemplate(dynamic list, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(list, tempPath, outPutPath);
        }
        public void GetFromTemplate(List<dynamic> lists, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(lists, tempPath, outPutPath);
        }
        public void GetFromTemplate<T>(List<T> list, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(list, tempPath, outPutPath);
        }
        public void GetFromTemplate<T>(List<List<T>> lists, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp<T>(lists, tempPath, outPutPath);
        }
    }
}
