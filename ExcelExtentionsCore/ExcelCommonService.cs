using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;

namespace ExcelExtentions.Core {
    internal class ExcelCommonService {
        public static ICollection<int> PicCols(DataTable dt) {
            ICollection<int> cs = new List<int>();
            int count = dt.Columns.Count;
            for (int i = 0; i < count; i++) {
                string colName = dt.Columns[i].ColumnName.ToLowerInvariant();
                if (IsPicture(colName))
                    cs.Add(i + 1);
            }
            return cs;
        }
        public static ICollection<int> PicCols(dynamic d) {
            ICollection<int> cs = new List<int>();
            var properties = PInfoBydynamics(d);
            int colCount = properties.Length;
            for (var i = 0; i < colCount; i++) {
                if (IsPicture(properties[i].Name))
                    cs.Add(i);
            }
            return cs;
        }
        public static PropertyInfo[] PInfoBydynamics(dynamic list) {
            Type type = null;
            foreach (var l in list) {
                type = l.GetType();
                break;
            }
            return type.GetProperties();
        }
        public static bool IsPicture(string str) {
            return str.Contains("img") || str.Contains("pic") || str.Contains("image");
        }
        public static void CreateFile(ExcelPackage ep, string fileName) {
            FileInfo newFile = new FileInfo(fileName);
            if (newFile.Exists) {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(fileName);
            }
            ep.SaveAs(newFile);
        }
        public static FileStream GetFileSteam(string path) {
            return File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        }
    }
}
