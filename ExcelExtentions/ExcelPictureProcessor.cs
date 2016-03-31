using OfficeOpenXml;
using System;
using System.Drawing;

namespace ExcelExtentions {
    internal class ExcelPictureProcessor {
        public static void SetPictureToCell(string picPath, ExcelWorksheet workSheet, int row, int col, int magicToken) {
            using (Bitmap img = new Bitmap(picPath)) {
                float hr = img.HorizontalResolution;
                //float percent = 1;
                //if (hr != 96) {
                //    percent = 96 / hr;
                //}
                //float picX = img.Width * percent;
                //float picY = img.Height * percent;
                float picX = img.Width;
                float picY = img.Height;

                double pWidth = picX * 2.02;
                double pHeight = picY * 0.755;

                double colWidth = workSheet.Column(col).Width * 15;
                double rowHeight = workSheet.Row(row).Height * 0.99;

                if (rowHeight < pHeight) {
                    rowHeight = pHeight / 0.9;
                    workSheet.Row(row).Height = rowHeight;
                }

                //I don't know why should double /2.....
                int offSetX = (int)((colWidth - pWidth) / 2 / 2);
                int offSetY = (int)((rowHeight - pHeight) / 2);
                offSetX = offSetX > 0 ? offSetX : 0;

                var pCell = workSheet.Drawings.AddPicture(Guid.NewGuid().ToString(), img, magicToken);
                //var pCell = workSheet.Drawings.AddPicture(Guid.NewGuid().ToString(), img);


                pCell.SetPosition(row - 1, offSetY, col - 1, offSetX);
                pCell.EditAs = OfficeOpenXml.Drawing.eEditAs.TwoCell;
            }
        }
    }
}
