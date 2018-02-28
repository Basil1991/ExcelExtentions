using System;

namespace ExcelExtentions.Core.Samples
{
    class Program
    {
        static void Main(string[] args)
        {
            try {
                //new ExcelHelpTests().GetByList();
                //new ExcelHelpTests().GetByLists();
                //new ExcelHelpTests().GetByDt();
                //new ExcelHelpTests().GetByDs();
                //new ExcelHelpTests().GetDynamic();
                //new ExcelHelpTests().GetByDynamicList();
                //new ExcelHelpTests().GetByDtFromTemp();
                //new ExcelHelpTests().GetByDsFromTemp();
                //new ExcelHelpTests().GetByDynamicFromTemp();
                //new ExcelHelpTests().GetByDynamicListFromTemp();
                new ExcelHelpTests().GetByListFromTemp();
                new ExcelHelpTests().GetByListsFromTemp();
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.Read();
            }
        }
    }
}
