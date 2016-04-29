using System.Collections.Generic;
using System.Data;
using ExcelExtentions.Argument;

namespace ExcelExtentions {
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
        public void Get(List<dynamic> list, ExcelArgument arg) {
            ExcelExport.Export(list, arg);
        }
        public void GetFromTemplate(DataTable dt, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(dt, tempPath, outPutPath);
        }
        public void GetFromTemplate(dynamic list, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(list, tempPath, outPutPath);
        }
        public void GetFromTemplate(List<dynamic> list, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(list, tempPath, outPutPath);
        }
        public void GetFromTemplate(DataSet ds, string tempPath, string outPutPath) {
            ExcelExportFromTemplate.ExportFromTemp(ds, tempPath, outPutPath);
        }
    }
}
