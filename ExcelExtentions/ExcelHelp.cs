using System.Collections.Generic;
using System.Data;
using ExcelExtentions.Argument;

namespace ExcelExtentions {
    public class ExcelHelp {
        public void Get(DataTable dt, ExcelArgument arg) {
            ExcelService.Export(dt, arg);
        }
        public void Get(DataSet ds, ExcelArgument arg) {
            ExcelService.Export(ds, arg);
        }
        public void Get(dynamic list, ExcelArgument arg) {
            ExcelService.Export(list, arg);
        }
        public void Get(List<dynamic> list, ExcelArgument arg) {
            ExcelService.Export(list, arg);
        }
    }
}
