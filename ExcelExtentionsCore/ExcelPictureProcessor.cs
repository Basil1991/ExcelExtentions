using OfficeOpenXml;
using System;
using System.Drawing;
using System.IO;

namespace ExcelExtentions.Core {
    internal class ExcelPictureProcessor {
        public static void SetPictureToCell(string picPath, ExcelWorksheet workSheet, int row, int col) {
            using (Bitmap img = new Bitmap(picPath)) {
                set(img, workSheet, row, col);
            }
        }
        public static void SetPictureToCell(Stream stream, ExcelWorksheet workSheet, int row, int col) {
            using (Bitmap img = new Bitmap(stream)) {
                set(img, workSheet, row, col);
            }
        }
        private static void set(Bitmap img, ExcelWorksheet workSheet, int row, int col) {
            float hr = img.HorizontalResolution;
            double hrPercent = 96 / hr;
            #region if DoAdjustDrawings=true.
            //float percent = 1;
            //if (hr != 96) {
            //    percent = 96 / hr;
            //}
            //float picX = img.Width * percent;
            //float picY = img.Height * percent;
            #endregion
            float picX = img.Width;
            float picY = img.Height;

            double pWidth = picX * 2.02;
            double pHeight = picY * 0.755;

            double picXDY = pWidth / pHeight;


            double colWidth = workSheet.Column(col).Width * 15;
            double rowHeight = workSheet.Row(row).Height * 0.99;

            int percent = 100;
            if (pWidth > colWidth * 0.85) {
                double pNWidth = colWidth * 0.85;
                percent = Convert.ToInt32(pNWidth / pWidth * 100 / hrPercent);
                pWidth = pNWidth;
                pHeight = pWidth / picXDY;
            }

            if (rowHeight < pHeight) {
                rowHeight = pHeight / 0.9;
                workSheet.Row(row).Height = rowHeight;
            }
            //I don't know why should double /2.....
            int offSetX = (int)((colWidth - pWidth) / 2 / 2);
            int offSetY = (int)((rowHeight - pHeight) / 2);
            offSetX = offSetX > 0 ? offSetX : 0;

            var pCell = workSheet.Drawings.AddPicture(Guid.NewGuid().ToString(), img);
            pCell.SetPosition(row - 1, offSetY, col - 1, offSetX);
            //
            if (percent != 100) {
                pCell.SetSize(percent);
            }
            pCell.EditAs = OfficeOpenXml.Drawing.eEditAs.TwoCell;
        }
    }
}
