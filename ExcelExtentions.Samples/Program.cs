using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtentions.Samples {
    class Program {
        static void Main(string[] args) {
            try {
                new ExcelHelpTests().GetByDt();
                new ExcelHelpTests().GetByDs();
                new ExcelHelpTests().GetDynamic();
                new ExcelHelpTests().GetByDynamicList();
                new ExcelHelpTests().GetByDtFromTemp();
                new ExcelHelpTests().GetByDsFromTemp();
                new ExcelHelpTests().GetByDynamicFromTemp();
                new ExcelHelpTests().GetByDynamicListFromTemp();
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.Read();
            }
        }
    }
}
