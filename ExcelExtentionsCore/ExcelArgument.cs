using System;
using System.Collections.Generic;

namespace ExcelExtentions.Core.Argument {
    public class ExcelArgument {
        public ExcelArgument(string outPutPath, List<SheetArgument> sheetArguments) {
            if (string.IsNullOrWhiteSpace(outPutPath))
                throw new ArgumentException("outPutPath");
            this.OutPutPath = outPutPath;
            if (sheetArguments == null || sheetArguments.Count <= 0) {
                throw new ArgumentException("sheetArguments");
            }
            this.SheetArguments = sheetArguments;
        }
        public string OutPutPath { get; private set; }
        public List<SheetArgument> SheetArguments { get; private set; }
    }
    public class SheetArgument {
        public SheetArgument(ColumnArgument[] columnArguments, string sheetName, short height = 20, bool isTitleShow = false, ClassType classType = ClassType.Default) {
            if (columnArguments == null) {
                throw new ArgumentException("columnArguments");
            }
            this.ColumnArguments = columnArguments;
            if (string.IsNullOrWhiteSpace(sheetName)) {
                this.SheetName = Guid.NewGuid().ToString();
            }
            else {
                this.SheetName = sheetName;
            }
            this.ClassType = classType;
            this.TitleHeight = 30;
            this.RowHeight = height;
        }
        public string ColumnHeight { get; private set; }
        public short TitleHeight { get; private set; }
        public short RowHeight { get; private set; }
        public ColumnArgument[] ColumnArguments { get; private set; }
        public ClassType ClassType { get; private set; }
        public string SheetName { get; private set; }
    }
    public enum ClassType {
        Default = 1,
        AllCenter = 2
    }
    public enum ColumnValueType {
        String,
        Int,
        DateTime,
        Date,
        Time,
        Double,
        Picture,
        IntNull,
        DoubleNull,
        Currency,
    }
    public class ColumnArgument {
        public ColumnArgument(ColumnValueType columnValueType, int width) {
            this.Width = width;
            this.ColumnValueType = columnValueType;
        }
        public ColumnArgument(ColumnValueType columnValueType) {
            this.Width = getDefaultWidthByColumnValueType(columnValueType);
            this.ColumnValueType = columnValueType;
        }
        public int Width { get; private set; }
        public ColumnValueType ColumnValueType { get; private set; }
        private int getDefaultWidthByColumnValueType(ColumnValueType type) {
            switch (type) {
                case ColumnValueType.Currency:
                case ColumnValueType.Date:
                case ColumnValueType.Time: return 12;
                case ColumnValueType.DateTime: return 20;
                case ColumnValueType.Double:
                case ColumnValueType.DoubleNull: return 10;
                case ColumnValueType.String: return 18;
                case ColumnValueType.Int:
                case ColumnValueType.IntNull: return 7;
                case ColumnValueType.Picture: return 25;
                default: return 10;
            }
        }
    }
}
