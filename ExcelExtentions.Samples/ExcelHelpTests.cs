using ExcelExtentions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using ExcelExtentions.Argument;
using System.Threading;

namespace ExcelExtentions.Samples {
    public class ExcelHelpTests {
        private string outPutDirPath = "../../OutPutDir/" + DateTime.Now.Millisecond.ToString();
        public static string PicPath = "../../Pictures/1.jpg";
        public void GetByDt() {
            var dt = getDt();

            Argument.ColumnArgument[] colArgs = getNormalColArgs();

            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet");
            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDT.xlsx"), sheetsArgs);
            new ExcelHelp().Get(dt, excelArgs);
        }
        public void GetByDs() {
            var ds = getDs();

            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet1");

            Argument.ColumnArgument[] colArgs2 = getNormalColArgs();
            Argument.SheetArgument sheetArgs2 = new Argument.SheetArgument(colArgs, "TestSheet2");

            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs, sheetArgs2 };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDS.xlsx"), sheetsArgs);
            new ExcelHelp().Get(ds, excelArgs);
        }
        public void GetDynamic() {
            var d = getDynamic();
            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument( colArgs, "TestSheet", classType: ClassType.AllCenter);
            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDynamic.xlsx"), sheetsArgs);
            new ExcelHelp().Get(d, excelArgs);
        }
        public void GetByDynamicList() {
            var ds = getDynamics();

            Argument.ColumnArgument[] colArgs = getNormalColArgs();
            Argument.SheetArgument sheetArgs = new Argument.SheetArgument(colArgs, "TestSheet1");

            Argument.ColumnArgument[] colArgs2 = getNormalColArgs();
            Argument.SheetArgument sheetArgs2 = new Argument.SheetArgument(colArgs, "TestSheet2");

            List<Argument.SheetArgument> sheetsArgs = new List<Argument.SheetArgument>() { sheetArgs, sheetArgs2 };
            Argument.ExcelArgument excelArgs = new Argument.ExcelArgument(string.Format(outPutDirPath + "_ByDynamics.xlsx"), sheetsArgs);
            new ExcelHelp().Get(ds, excelArgs);
        }
        private Argument.ColumnArgument[] getNormalColArgs() {
            Argument.ColumnArgument[] colArgs = new Argument.ColumnArgument[] {
            new Argument.ColumnArgument(Argument.ColumnValueType.Int),
            new Argument.ColumnArgument(Argument.ColumnValueType.String),
            new Argument.ColumnArgument(Argument.ColumnValueType.DateTime),
            new Argument.ColumnArgument(Argument.ColumnValueType.Double),
            new Argument.ColumnArgument(Argument.ColumnValueType.Picture),
            new Argument.ColumnArgument(Argument.ColumnValueType.Currency),
            };
            return colArgs;
        }
        private DataTable getDt() {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Text");
            dt.Columns.Add("Datetime");
            dt.Columns.Add("DoubleValue");
            dt.Columns.Add("Pictures");
            dt.Columns.Add("Money");

            for (int i = 0; i < 1 * 10; i++) {
                DataRow nRow = dt.NewRow();
                nRow["ID"] = i;
                nRow["Text"] = "123123123" + i;
                nRow["Datetime"] = DateTime.Now.AddDays(i);
                nRow["DoubleValue"] = new Random().NextDouble();
                if (i % 2 == 0) {
                    int imgSeed = new Random().Next(1, 10);
                    Thread.Sleep(10);
                    nRow["Pictures"] = string.Format("../../Pictures/{0}.jpg", imgSeed);
                }
                else {
                    //nRow["Pictures"] = "http://www.52ij.com/uploads/allimg/160317/1110104P8-4.jpg";
                    nRow["Pictures"] = string.Format("../../Pictures/9.png");
                }
                nRow["Money"] = "CAD 12.11";
                dt.Rows.Add(nRow);
            }
            ExcelHelp eh = new ExcelHelp();
            return dt;
        }
        private DataSet getDs() {
            DataTable dt = getDt();
            DataTable dt1 = getDt();
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            ds.Tables.Add(dt1);
            return ds;
        }
        private dynamic getDynamic() {
            List<User> user = new List<User>();
            for (int i = 0; i < 100; ++i) {
                user.Add(new User(true));
            }
            dynamic d = user.Select(a => new {
                Age = a.Age,
                Name = a.Name,
                BirthDay = a.BirthDate,
                Height = a.Height,
                Pic = a.PicturePath,
                Money = "USD 123.22"
            }).ToList();

            return d;
        }
        private List<dynamic> getDynamics() {
            List<dynamic> dList = new List<dynamic>();
            dList.Add(getDynamic());
            dList.Add(getDynamic());
            return dList;
        }
    }
    public class User {
        public User() {
        }
        public User(bool isDefalt) {
            if (!isDefalt) { }
            else {
                Name = "Lilei" + new Random().Next(1, 10000);
                Age = new Random().Next(10, 50);
                Height = 182.25;
                BirthDate = DateTime.Now.AddDays(0 - new Random().Next(365 * 10, 365 * 100));
                PicturePath = ExcelHelpTests.PicPath;
                Thread.Sleep(10);
            }
        }
        public string Name { get; set; }
        public int Age { get; set; }
        public string PicturePath { get; set; }
        public double Height { get; set; }
        public DateTime BirthDate { get; set; }
    }
}